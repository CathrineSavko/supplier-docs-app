import argparse
import copy
import hashlib
import json
import re
import shutil
import sys
from decimal import Decimal, InvalidOperation
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

from process_commercial_offer import normalize, update_progress


W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS = {'w': W, 'r': R, 'pr': P}
QN = lambda name: f'{{{W}}}{name}'
RQN = lambda name: f'{{{R}}}{name}'
INVOICE = re.compile(r'(?<![A-Z0-9])([A-Z]{2,10}\d+[A-Z]*)(?=[^A-Z0-9]|[a-z]|$)')
AMOUNT = re.compile(r'amount\s+of\s+CNY\s*([0-9][0-9\s.,]*)', re.I)


def file_hash(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def read_xml(docx_path):
    with ZipFile(docx_path) as package:
        return etree.fromstring(package.read('word/document.xml'))


def body_children(root):
    body = root.find('w:body', NS)
    if body is None:
        raise ValueError('Word document body is missing.')
    return body, list(body)


def has_section_break(element):
    return bool(element.xpath('.//w:sectPr', namespaces=NS))


def document_sections(root):
    body, children = body_children(root)
    sections, current = [], []
    final_section = None
    for child in children:
        if child.tag == QN('sectPr'):
            final_section = child
            continue
        current.append(copy.deepcopy(child))
        if has_section_break(child):
            sections.append(current)
            current = []
    if current:
        if final_section is None:
            raise ValueError('The final Word section properties are missing.')
        current.append(copy.deepcopy(final_section))
        sections.append(current)
    return sections


def section_text(section):
    return ''.join(''.join(element.xpath('.//w:t/text()', namespaces=NS)) for element in section)


def decimal_key(value):
    candidate = re.sub(r'[^0-9,.-]', '', str(value)).replace(',', '.').strip('.,')
    try:
        return format(Decimal(candidate).normalize(), 'f')
    except InvalidOperation as error:
        raise ValueError(f'Invalid payment amount: {value}') from error


def amount_for_translation(value):
    numeric = Decimal(decimal_key(value))
    text = f'{numeric:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', ' ')
    return text


def details_from_section(section, page_number):
    text = section_text(section)
    matched = AMOUNT.search(text)
    if not matched:
        raise ValueError(f'Payment amount was not found on letter page {page_number}.')
    invoices = INVOICE.findall(text)
    if not invoices:
        raise ValueError(f'Invoice numbers were not found on letter page {page_number}.')
    return {'page': page_number, 'amount': decimal_key(matched.group(1)), 'invoices': invoices, 'source_text': text}


def containers_from_invoices(invoices):
    """Return the three-digit container part from CTI25xxxBY-style invoice numbers."""
    containers = []
    for invoice in invoices:
        matched = re.search(r'(\d{3})(?=BY$)', invoice, re.I)
        if not matched:
            raise ValueError(f'Cannot determine a container number from invoice {invoice}.')
        container = matched.group(1)
        if container not in containers:
            containers.append(container)
    if not containers:
        raise ValueError('Container numbers were not found in the supplier letter.')
    return ' '.join(containers)


def invoice_list(invoices):
    if len(invoices) == 1:
        return invoices[0]
    if len(invoices) == 2:
        return f'{invoices[0]} и {invoices[1]}'
    return f'{", ".join(invoices[:-1])} и {invoices[-1]}'


def paragraph_text(element):
    return ''.join(element.xpath('.//w:t/text()', namespaces=NS))


def replace_text_span(paragraph, pattern, replacement):
    text_nodes = paragraph.xpath('.//w:t', namespaces=NS)
    text = ''.join(node.text or '' for node in text_nodes)
    matched = re.search(pattern, text, re.S)
    if not matched:
        return False
    start, end = matched.span()
    cursor, inserted = 0, False
    for node in text_nodes:
        value = node.text or ''
        node_start, node_end = cursor, cursor + len(value)
        cursor = node_end
        if node_end <= start or node_start >= end:
            continue
        prefix = value[:max(0, start - node_start)] if node_start <= start else ''
        suffix = value[max(0, end - node_start):] if node_start <= end <= node_end else ''
        if not inserted:
            node.text = prefix + replacement + suffix
            inserted = True
        else:
            node.text = suffix
    return True


def translation_elements(template_root, details, translation_date):
    _, children = body_children(template_root)
    elements = [copy.deepcopy(element) for element in children if element.tag != QN('sectPr')]
    for element in elements:
        if element.tag != QN('p'):
            continue
        text = paragraph_text(element)
        if 'полная оплата' in text:
            replace_text_span(element, r'.*', f'Настоящим уведомляем Вас о том, что нами получена полная оплата в размере {amount_for_translation(details["amount"])} юаня (CNY).')
        elif 'счетам-фактурам' in text:
            replace_text_span(element, r'.*', f'По счетам-фактурам (инвойсам) №: {invoice_list(details["invoices"])} от {translation_date} г. в рамках Договора № 101-L от 06.06.2024.')
        elif 'Перевод верен' in text or text.startswith('Дата:'):
            replace_text_span(element, r'Дата:\s*.*?(?=Перевод верен)', f'Дата: {translation_date} г.')
            replace_text_span(element, r'Перевод верен\.\s*.*?\.', f'Перевод верен. {translation_date}.')
    return elements, copy.deepcopy(children[-1])


def page_elements_for_output(letter_elements):
    elements = [copy.deepcopy(element) for element in letter_elements]
    if elements and elements[-1].tag == QN('sectPr'):
        section_properties = elements.pop()
        last_paragraph = elements[-1]
        paragraph_properties = last_paragraph.find('w:pPr', NS)
        if paragraph_properties is None:
            paragraph_properties = etree.Element(QN('pPr'))
            last_paragraph.insert(0, paragraph_properties)
        paragraph_properties.append(section_properties)
    return elements


def source_content_signature(section):
    copied = [copy.deepcopy(element) for element in section]
    visible = []
    for element in copied:
        if element.tag == QN('sectPr'):
            continue
        for section_properties in element.xpath('.//w:sectPr', namespaces=NS):
            section_properties.getparent().remove(section_properties)
        visible.append(element)
    return hashlib.sha256(b''.join(etree.tostring(element, with_tail=False) for element in visible)).digest()


def merge_template_relationships(source_package, template_package, translation_elements):
    source_rels = etree.fromstring(source_package.read('word/_rels/document.xml.rels'))
    template_rels = etree.fromstring(template_package.read('word/_rels/document.xml.rels'))
    referenced = {
        value
        for top_level in translation_elements
        for element in top_level.iter()
        for name, value in element.attrib.items()
        if name.startswith(f'{{{R}}}')
    }
    template_by_id = {item.get('Id'): item for item in template_rels.findall('pr:Relationship', NS)}
    used_numbers = [int(match.group(1)) for item in source_rels.findall('pr:Relationship', NS) if (match := re.fullmatch(r'rId(\d+)', item.get('Id', '')))]
    next_number = max(used_numbers, default=0) + 1
    replacements, extra_parts = {}, {}
    for old_id in referenced:
        relation = template_by_id.get(old_id)
        if relation is None:
            continue
        new_id = f'rId{next_number}'
        next_number += 1
        new_relation = copy.deepcopy(relation)
        new_relation.set('Id', new_id)
        target = new_relation.get('Target', '')
        if new_relation.get('TargetMode') != 'External' and target:
            source_part = str(Path('word') / target).replace('\\', '/')
            source_part = str(Path(source_part)).replace('\\', '/')
            if source_part in template_package.namelist():
                destination_target = f'media/translation-{Path(target).name}'
                new_relation.set('Target', destination_target)
                extra_parts[f'word/{destination_target}'] = template_package.read(source_part)
        source_rels.append(new_relation)
        replacements[old_id] = new_id
    for top_level in translation_elements:
        for element in top_level.iter():
            for name, value in list(element.attrib.items()):
                if name.startswith(f'{{{R}}}') and value in replacements:
                    element.set(name, replacements[value])
    return etree.tostring(source_rels, xml_declaration=True, encoding='UTF-8', standalone=True), extra_parts


def write_output(source_path, template_path, destination, source_root, letter_elements, template_root, details, translation_date):
    output_root = copy.deepcopy(source_root)
    output_body, _ = body_children(output_root)
    for child in list(output_body):
        output_body.remove(child)
    for element in page_elements_for_output(letter_elements):
        output_body.append(element)
    translated, final_section = translation_elements(template_root, details, translation_date)
    for element in translated:
        output_body.append(element)
    output_body.append(final_section)
    with ZipFile(source_path) as source, ZipFile(template_path) as template, ZipFile(destination, 'w', ZIP_DEFLATED) as target:
        rels, extra_parts = merge_template_relationships(source, template, translated)
        xml = etree.tostring(output_root, xml_declaration=True, encoding='UTF-8', standalone=True)
        for info in source.infolist():
            if info.filename == 'word/document.xml':
                target.writestr(info, xml)
            elif info.filename == 'word/_rels/document.xml.rels':
                target.writestr(info, rels)
            else:
                target.writestr(info, source.read(info.filename))
        for name, data in extra_parts.items():
            target.writestr(name, data)


def find_translation_template(templates_dir):
    templates = list((Path(templates_dir) / 'translations').glob('*.docx'))
    if len(templates) != 1:
        raise ValueError('Translation page template is missing or ambiguous.')
    return templates[0]


def safe_filename(value):
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', value).strip()[:170]


def validate_output(path, source_first_section, details, translation_date):
    output_root = read_xml(path)
    sections = document_sections(output_root)
    if len(sections) != 2:
        raise ValueError('Validation failed: the output must contain exactly two pages/sections.')
    if source_content_signature(source_first_section) != source_content_signature(sections[0]):
        raise ValueError('Validation failed: the original supplier letter page changed.')
    translated = section_text(sections[1])
    if amount_for_translation(details['amount']) not in translated:
        raise ValueError('Validation failed: translated payment amount is incorrect.')
    if any(invoice not in translated for invoice in details['invoices']):
        raise ValueError('Validation failed: translated invoice numbers are incomplete.')
    if translated.count(translation_date) < 2:
        raise ValueError('Validation failed: translated date is incorrect.')


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--job-dir', required=True)
    parser.add_argument('--templates-dir', required=True)
    args = parser.parse_args()
    job_dir = Path(args.job_dir)
    request = json.loads((job_dir / 'request.json').read_text(encoding='utf-8'))
    params = request.get('parameters', {})
    translation_date = normalize(params.get('translationDate'))
    if not translation_date:
        raise ValueError('Translation date is required.')
    sources = list((job_dir / 'input').glob('*.docx'))
    if len(sources) != 1:
        raise ValueError('Task 5 needs exactly one multi-page supplier .docx file.')
    source = sources[0]
    template = find_translation_template(args.templates_dir)
    source_hash = file_hash(source)
    template_hash = file_hash(template)
    update_progress(job_dir, 8, 'Reading supplier letters...')
    source_root = read_xml(source)
    sections = document_sections(source_root)
    details = [details_from_section(section, index) for index, section in enumerate(sections, start=1)]
    selected = []
    for item in details:
        selected.append((item, containers_from_invoices(item['invoices'])))
    update_progress(job_dir, 28, f'Found supplier letters: {len(selected)}')
    template_root = read_xml(template)
    output_dir = job_dir / 'output'
    output_dir.mkdir(exist_ok=True)
    output_files = []
    for index, (item, containers) in enumerate(selected, start=1):
        update_progress(job_dir, 30 + int(50 * index / max(1, len(selected))), f'Creating letter {index} of {len(selected)}...')
        name = safe_filename(f'Информационное письмо {containers}.docx')
        destination = output_dir / name
        write_output(source, template, destination, source_root, sections[item['page'] - 1], template_root, item, translation_date)
        validate_output(destination, sections[item['page'] - 1], item, translation_date)
        output_files.append(name)
    if file_hash(source) != source_hash or file_hash(template) != template_hash:
        raise ValueError('A retained source or translation template changed unexpectedly.')
    print(json.dumps({'outputFiles': output_files, 'notes': [f'Letters created: {len(output_files)}; originals and mapped values verified.']}, ensure_ascii=True))


if __name__ == '__main__':
    try:
        main()
    except Exception as error:
        print(str(error), file=sys.stderr)
        sys.exit(1)

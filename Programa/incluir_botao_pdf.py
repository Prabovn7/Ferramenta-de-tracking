"""Acrescenta ao XLSM o botão e o projeto VBA criados e testados no Excel.

Não executa macros, não usa COM e não altera a configuração de segurança.
"""
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET

S='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P='http://schemas.openxmlformats.org/package/2006/relationships'
C='http://schemas.openxmlformats.org/package/2006/content-types'


def xml_bytes(root, namespace):
    # OPC exige o namespace padrão em Types/Relationships em alguns leitores.
    ET.register_namespace('',namespace)
    ET.register_namespace('r',R)
    return ET.tostring(root,encoding='utf-8',xml_declaration=True)


def incluir_botao(entrada: Path, saida: Path):
    assets=Path(__file__).resolve().parent/'recursos'
    vba=(assets/'vbaProject.bin').read_bytes()
    desenho=(assets/'botao_pdf.xml').read_bytes()
    with ZipFile(entrada) as z:
        parts={n:z.read(n) for n in z.namelist()}
    if 'xl/drawings/drawing1.xml' in parts:
        raise ValueError('O painel já possui um desenho. Não foi possível inserir o botão sem substituir conteúdo.')
    # O projeto usa nomes de planilha e contém apenas ações invocadas pelo usuário.
    wb=ET.fromstring(parts['xl/workbook.xml'])
    pr=wb.find(f'{{{S}}}workbookPr')
    if pr is None:pr=ET.SubElement(wb,f'{{{S}}}workbookPr')
    pr.set('codeName','EstaPastaDeTrabalho')
    parts['xl/workbook.xml']=xml_bytes(wb,S)
    for i in range(1,4):
        name=f'xl/worksheets/sheet{i}.xml';sheet=ET.fromstring(parts[name])
        pr=sheet.find(f'{{{S}}}sheetPr')
        if pr is None:
            pr=ET.Element(f'{{{S}}}sheetPr');sheet.insert(0,pr)
        pr.set('codeName',f'Planilha{i}')
        if i==1:
            relname='xl/worksheets/_rels/sheet1.xml.rels'
            rels=ET.fromstring(parts[relname]) if relname in parts else ET.Element(f'{{{P}}}Relationships')
            ids={x.get('Id') for x in rels}
            rid='rIdPDF'
            if rid in ids:raise ValueError('Relacionamento do botão já existe.')
            ET.SubElement(rels,f'{{{P}}}Relationship',Id=rid,Type=R+'/drawing',Target='../drawings/drawing1.xml')
            parts[relname]=xml_bytes(rels,P)
            ET.SubElement(sheet,f'{{{S}}}drawing',{f'{{{R}}}id':rid})
        parts[name]=xml_bytes(sheet,S)
    rels=ET.fromstring(parts['xl/_rels/workbook.xml.rels'])
    ET.SubElement(rels,f'{{{P}}}Relationship',Id='rIdVBA',Type='http://schemas.microsoft.com/office/2006/relationships/vbaProject',Target='vbaProject.bin')
    parts['xl/_rels/workbook.xml.rels']=xml_bytes(rels,P)
    types=ET.fromstring(parts['[Content_Types].xml'])
    for item in types:
        if item.get('PartName')=='/xl/workbook.xml':
            item.set('ContentType','application/vnd.ms-excel.sheet.macroEnabled.main+xml')
    ET.SubElement(types,f'{{{C}}}Override',PartName='/xl/vbaProject.bin',ContentType='application/vnd.ms-office.vbaProject')
    ET.SubElement(types,f'{{{C}}}Override',PartName='/xl/drawings/drawing1.xml',ContentType='application/vnd.openxmlformats-officedocument.drawing+xml')
    parts['[Content_Types].xml']=xml_bytes(types,C)
    parts['xl/vbaProject.bin']=vba
    parts['xl/drawings/drawing1.xml']=desenho
    with ZipFile(saida,'w',ZIP_DEFLATED) as z:
        for name,content in parts.items():z.writestr(name,content)

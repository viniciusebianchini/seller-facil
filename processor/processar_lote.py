"""Gera etiquetas Shopee, ZIP, Word e conferência em uma pasta nova.

Uso: python processar_lote.py pedidos.xlsx etiquetas.zip pasta_nova --previews pasta_de_revisao
Dependências: Pillow, openpyxl, reportlab, python-docx. Sem conexão externa.
"""
import argparse
import csv
import json
import re
import tempfile
from pathlib import Path
from collections import Counter
from datetime import datetime

import openpyxl
from PIL import Image, ImageOps
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import etiquetas


def read_products(path):
    wb=openpyxl.load_workbook(path,data_only=True,read_only=True)
    try:
        wb.active.reset_dimensions()
        values=iter(wb.active.values)
        headers=next(values)
        result={}
        for values_row in values:
            if not any(v is not None for v in values_row):continue
            row=dict(zip(headers,values_row))
            tracking=row['tracking_number']
            etiquetas.require(tracking not in result,'Rastreio duplicado.')
            row['items']=etiquetas.items_from(row['product_info'])
            result[tracking]=row
        return result
    finally:wb.close()


def add_table(doc,headers,records,widths,numeric):
    table=doc.add_table(rows=1,cols=len(headers))
    table.autofit=False
    for col,width in zip(table.columns,widths):col.width=Cm(width)
    for cell,value in zip(table.rows[0].cells,headers):cell.text=value
    table.rows[0]._tr.get_or_add_trPr().append(OxmlElement('w:tblHeader'))
    for record in records:
        for cell,value in zip(table.add_row().cells,record):cell.text=str(value)
    for index,row in enumerate(table.rows):
        row._tr.get_or_add_trPr().append(OxmlElement('w:cantSplit'))
        for column,(cell,width) in enumerate(zip(row.cells,widths)):
            cell.width=Cm(width)
            cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            props=cell._tc.get_or_add_tcPr()
            borders=OxmlElement('w:tcBorders')
            for side in ('top','bottom','left','right'):
                edge=OxmlElement('w:'+side)
                for key,value in [('val','single'),('sz','4'),('color','D9D9D9')]:edge.set(qn('w:'+key),value)
                borders.append(edge)
            props.append(borders)
            margins=OxmlElement('w:tcMar')
            for side,value in [('top','75'),('bottom','75'),('left','90'),('right','90')]:
                edge=OxmlElement('w:'+side);edge.set(qn('w:w'),value);edge.set(qn('w:type'),'dxa');margins.append(edge)
            props.append(margins)
            if index==0:
                shade=OxmlElement('w:shd');shade.set(qn('w:fill'),'E7E6E6');props.append(shade)
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.space_after=Pt(0)
                paragraph.paragraph_format.line_spacing=1
                paragraph.alignment=WD_ALIGN_PARAGRAPH.CENTER if column in numeric else WD_ALIGN_PARAGRAPH.LEFT
                for run in paragraph.runs:
                    run.font.size=Pt(10.5)
                    run.bold=index==0 or column in numeric
    return table


def write_word(xlsx,order_file,path,issued_at=None):
    issued_at=issued_at or datetime.now().astimezone()
    products=read_products(xlsx)
    with Path(order_file).open(encoding='utf-8-sig',newline='') as source:
        sequence=list(csv.DictReader(source,delimiter=';'))
    etiquetas.require(len(sequence)==len(products) and {r['Rastreio'] for r in sequence}==set(products),'A lista e a ordem de impressao nao correspondem.')
    counts=Counter();descriptions={}
    for row in products.values():
        for item in row['items']:
            sku=item['sku'];counts[sku]+=item['qty']
            desc=item['name']+('\nVariação: '+item['variation'] if item['variation'] else '')
            descriptions.setdefault(sku,set()).add(desc)
    doc=Document()
    section=doc.sections[0]
    section.page_width=Cm(21);section.page_height=Cm(29.7)
    section.top_margin=section.bottom_margin=Cm(1.7)
    section.left_margin=section.right_margin=Cm(1.6)
    for name in ('Normal','Title','Heading 1','Heading 2'):
        style=doc.styles[name]
        style.font.name='Arial';style.font.color.rgb=RGBColor(0,0,0)
    doc.styles['Normal'].font.size=Pt(10.5)
    doc.styles['Normal'].paragraph_format.space_after=Pt(6)
    doc.styles['Title'].font.size=Pt(20)
    doc.add_paragraph('Separação de produtos por SKU','Title')
    doc.add_paragraph('Emissão: '+issued_at.strftime('%d/%m/%Y às %H:%M:%S')+' (horário local)')
    doc.add_paragraph(f"Shopee • {len({r['order_sn'] for r in products.values()})} pedidos • {sum(counts.values())} unidades de venda • {len(counts)} SKUs")
    doc.add_paragraph('Separar as quantidades abaixo e conferir o SKU e a variação. Kits e conjuntos contam como unidades de venda, com sua composição completa.')
    doc.add_paragraph('Responsável: ________________________    Data: ____/____/________')
    records=[[sku,' / '.join(sorted(descriptions[sku])),counts[sku],''] for sku in sorted(counts,key=str.casefold)]
    records.append(['TOTAL',f'{len(counts)} SKUs',sum(counts.values()),''])
    add_table(doc,['SKU','Produto e variação','Qtd.','Conf.'],records,[5.1,10.1,1.2,1.4],{2,3})
    doc.add_paragraph('Pedidos na ordem de impressão','Heading 1')
    order_records=[]
    for record in sequence:
        row=products[record['Rastreio']]
        labels=record['Primeira etiqueta']
        if labels!=record['Ultima etiqueta']:labels+=' a '+record['Ultima etiqueta']
        detail='\n'.join(f"{item['sku']} | Qtd: {item['qty']}" for item in row['items'])
        order_records.append([labels,row['order_sn']+'\n'+row['tracking_number'],detail,''])
    add_table(doc,['Etiqueta','Pedido e rastreio','SKUs e quantidades','Conf.'],order_records,[1.6,7.2,7.6,1.4],{0,3})
    doc.add_paragraph('Fonte: '+Path(xlsx).name)
    footer=section.footer.paragraphs[0];footer.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    footer.add_run('Página ')
    field=OxmlElement('w:fldSimple');field.set(qn('w:instr'),'PAGE');footer._p.append(field)
    doc.save(path)
    check=Document(path)
    etiquetas.require(sum(int(r.cells[2].text) for r in check.tables[0].rows[1:-1])==sum(counts.values()),'Total do Word inconsistente.')
    etiquetas.require(len(check.tables[1].rows)-1==len(sequence),'Pedidos ausentes no Word.')
    return {'skus':len(counts),'unidades':sum(counts.values())}


def previews(zpl,directory):
    directory=Path(directory);directory.mkdir(parents=True,exist_ok=True)
    matches=re.findall(r'~DG[^,]+,(\d+),(\d+),([0-9A-F]+)\^XA',Path(zpl).read_text(encoding='ascii'))
    for number,(size,stride,data) in enumerate(matches,1):
        raw=bytes.fromhex(data)
        etiquetas.require(len(raw)==int(size),'Imagem ZPL incompleta.')
        im=ImageOps.invert(Image.frombytes('1',(int(stride)*8,int(size)//int(stride)),raw).convert('L'))
        im.save(directory/f'etiqueta-{number:03}.png')
    return len(matches)


def write_pdf(zpl,path,expected_pages):
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import mm
    from reportlab.lib.utils import ImageReader
    from pypdf import PdfReader
    matches=re.findall(r'~DG[^,]+,(\d+),(\d+),([0-9A-F]+)\^XA',Path(zpl).read_text(encoding='ascii'))
    etiquetas.require(len(matches)==expected_pages,'Quantidade de paginas PDF inconsistente.')
    pdf=canvas.Canvas(str(path),pagesize=(100*mm,150*mm),pageCompression=1)
    pdf.setTitle('Etiquetas Shopee')
    for size,stride,data in matches:
        raw=bytes.fromhex(data)
        etiquetas.require(len(raw)==int(size),'Imagem incompleta para PDF.')
        im=ImageOps.invert(Image.frombytes('1',(int(stride)*8,int(size)//int(stride)),raw).convert('L'))
        scale=min(100*mm/im.width,150*mm/im.height)
        width,height=im.width*scale,im.height*scale
        pdf.drawImage(ImageReader(im),(100*mm-width)/2,(150*mm-height)/2,width=width,height=height)
        pdf.showPage()
    pdf.save()
    check=PdfReader(path)
    etiquetas.require(len(check.pages)==expected_pages,'PDF incompleto.')
    etiquetas.require(all(abs(float(page.mediabox.width)-100*mm)<0.01 and abs(float(page.mediabox.height)-150*mm)<0.01 for page in check.pages),'Tamanho de pagina PDF incorreto.')


def process(xlsx,archive,destination,preview_dir=None):
    destination=Path(destination).resolve()
    etiquetas.require(not destination.exists(),'Escolha uma pasta de saida nova para preservar os lotes anteriores.')
    destination.parent.mkdir(parents=True,exist_ok=True)
    # Publish the entire batch only after all matching and document checks succeed.
    with tempfile.TemporaryDirectory(prefix='shopee-',dir=destination.parent) as temporary:
        stage=Path(temporary)/'lote'
        result=etiquetas.run(xlsx,archive,stage)
        zpl=Path(result['arquivo'])
        pdf=zpl.with_suffix('.pdf')
        write_pdf(zpl,pdf,result['etiquetas'])
        word=stage/(zpl.stem.replace('etiquetas-com-separacao','lista-de-separacao')+'.docx')
        stamp=re.search(r' - (\d{4}-\d{2}-\d{2} \d{2}-\d{2}-\d{2}) - ',zpl.stem)
        etiquetas.require(stamp is not None,'Data de emissao ausente no nome do lote.')
        issued_at=datetime.strptime(stamp[1],'%Y-%m-%d %H-%M-%S')
        result.update(write_word(xlsx,stage/'ordem-de-impressao.csv',word,issued_at=issued_at))
        if preview_dir:
            etiquetas.require(previews(zpl,preview_dir)==result['etiquetas'],'Quantidade de previas inconsistente.')
        result['arquivo']=str(destination/zpl.name)
        result['zip']=str(destination/zpl.with_suffix('.zip').name)
        result['word']=str(destination/word.name)
        result['pdf']=str(destination/pdf.name)
        (stage/'resumo.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
        stage.rename(destination)
    return result


if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('planilha');parser.add_argument('zip_original');parser.add_argument('pasta_saida')
    parser.add_argument('--previews',help='Pasta de imagens para revisão; não enviar como resultado padrão.')
    args=parser.parse_args()
    print(json.dumps(process(args.planilha,args.zip_original,args.pasta_saida,args.previews),ensure_ascii=False))

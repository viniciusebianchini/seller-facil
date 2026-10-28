"""Shopee: agrupa pedidos por SKU e mantem envio + separacao consecutivos.
Requer Python 3, Pillow, openpyxl e reportlab. Layout suportado: 816 x 1218 pontos.
Uso: python gerar_etiquetas.py "pedidos.xlsx" "etiquetas.zip" "pasta_saida"
"""
import argparse,base64,collections,csv,re,textwrap,zipfile,zlib
from datetime import datetime
from pathlib import Path
from PIL import Image,ImageOps,ImageDraw,ImageFont,ImageChops
from reportlab.graphics.barcode.code128 import _patterns
import openpyxl
patterns={tuple(ord(c.lower())-96 for c in v):k for k,v in _patterns.items()}
def barcode(im,y):
    a=[im.getpixel((x,y))<128 for x in range(im.width)]
    start=next((i for i,v in enumerate(a) if v),None)
    if start is None:return None
    runs=[]; prev=True; n=0
    for v in a[start:]:
        if v==prev:n+=1
        else:runs.append(n);n=1;prev=v
    runs.append(n)
    codes=[]
    for i in range(0,len(runs)-6,6):
        r=runs[i:i+6]; unit=sum(r)/11
        key=tuple(round(v/unit) for v in r)
        if key not in patterns:break
        codes.append(patterns[key])
    if len(codes)<4 or codes[0] not in (103,104,105):return None
    # Stop pattern uses 13 modules and seven runs.
    j=6*len(codes); stop=runs[j:j+7]
    if len(stop)!=7 or tuple(round(v/(sum(stop)/13)) for v in stop)!=(2,3,3,1,1,1,2):return None
    if (codes[0]+sum(i*v for i,v in enumerate(codes[1:-1],1)))%103!=codes[-1]:return None
    mode=codes[0]-103; out=''
    for c in codes[1:-1]:
        if c==99 and mode!=2:mode=2
        elif c==100:mode=1
        elif c==101:mode=0
        elif mode==2:out+=f'{c:02}'
        elif c<96:out+=chr(c+32 if mode==1 or c<64 else c-64)
        else:return None
    return out

ITEM=re.compile(r'\[\d+\] Product Name:(.*?); Variation Name:(.*?); Price:(.*?); Quantity:\s*(.*?); SKU Reference No\.:\s*(.*?); Parent SKU Reference No\.:\s*(.*?);',re.S)
GRAPHIC=re.compile(r'~DG([^,]+),(\d+),(\d+),:Z64:([^:]+):([0-9A-Fa-f]+)')

def require(ok,message):
    if not ok:raise ValueError(message)

def font(size,bold=False):
    for name in (["C:/Windows/Fonts/arialbd.ttf","DejaVuSans-Bold.ttf"] if bold else ["C:/Windows/Fonts/arial.ttf","DejaVuSans.ttf"]):
        try:return ImageFont.truetype(name,size)
        except OSError:pass
    raise ValueError('Fonte Arial ou DejaVu Sans nao encontrada.')

def wrap(value,f,width=780):
    lines=['']
    for word in value.split():
        candidate=(lines[-1]+' '+word).strip()
        if f.getlength(candidate)<=width:lines[-1]=candidate
        else:
            if lines[-1]:lines.append('')
            for char in word:
                if f.getlength(lines[-1]+char)>width:lines.append('')
                lines[-1]+=char
    return lines

def items_from(info):
    require(isinstance(info,str) and info.strip(),'Lista de produtos ausente.')
    info=info.replace('\\r\\n','\n').replace('\\n','\n')
    matches=list(ITEM.finditer(info))
    require(matches and len(matches)==len(re.findall(r'\[\d+\]\s*Product Name:',info)), 'Lista de produtos incompleta ou formato desconhecido.')
    require(not ITEM.sub('',info).strip(),'Conteudo de produto nao reconhecido.')
    items=[]
    for m in matches:
        name,variation,price,qty,sku,parent=(v.strip() for v in m.groups())
        require(name and sku and qty.isdigit() and int(qty)>0,'Produto ou SKU ausente ou quantidade invalida.')
        items.append(dict(name=name,variation=variation,sku=sku,qty=int(qty)))
    return items

def lines_for(items,large=False):
    heading=font(25 if large else 23,True)
    detail=font(21 if large else 17)
    lines=[]
    for item in items:
        lines.extend((line,heading,31 if large else 27) for line in wrap(f"SKU: {item['sku']} | Qtd: {item['qty']}",heading))
        description=item['name']
        if item['variation']:description+=' | '+item['variation']
        lines.extend((line,detail,26 if large else 20) for line in wrap(description,detail))
        if large:lines.append(('',detail,12))
    return lines

def paint(im,lines,y):
    d=ImageDraw.Draw(im)
    for line,f,step in lines:
        d.text((18,y),line,font=f,fill=0,anchor='lt');y+=step
    return y

def graphic(im,name='R:DEMO.GRF'):
    raw=ImageOps.invert(im.convert('L')).point(lambda v:255 if v>=96 else 0).convert('1').tobytes()
    return f'~DG{name},{len(raw)},102,'+raw.hex().upper()

def separation_pages(row,items):
    # Keep every item's heading and description together when it fits on one page.
    pages=[]; current=[]; used=0; capacity=960
    for item in items:
        lines=lines_for([item],True); height=sum(x[2] for x in lines)
        if current and used+height>capacity:pages.append(current);current=[];used=0
        for line in lines:
            if used+line[2]>capacity:pages.append(current);current=[];used=0
            current.append(line);used+=line[2]
    if current:pages.append(current)
    result=[]; previews=[]
    for i,lines in enumerate(pages,1):
        im=Image.new('L',(816,1218),255);d=ImageDraw.Draw(im)
        d.text((18,24),'SEPARACAO DO PEDIDO',font=font(34,True),fill=0)
        d.text((18,82),'Pedido: '+row['order_sn'],font=font(28,True),fill=0)
        d.text((18,125),'Rastreio: '+row['tracking_number'],font=font(23),fill=0)
        d.text((18,166),f'Pagina {i}/{len(pages)} | Acompanha a etiqueta de envio anterior',font=font(22),fill=0)
        d.line((18,208,798,208),fill=0,width=2)
        paint(im,lines,225)
        result.append(graphic(im)+'^XA^MMT,Y^PON^MNY^FO0,0^XGR:DEMO.GRF,1,1^FS^PQ1,0,0,N^XZ^XA^IDR:DEMO.GRF^FS^XZ\n')
        previews.append(im)
    return result,previews

def quantity_alert(row,items):
    total=sum(item['qty'] for item in items)
    im=Image.new('L',(816,1218),255)
    d=ImageDraw.Draw(im)
    d.rectangle((18,18,797,1199),outline=0,width=7)
    d.rectangle((30,30,785,190),fill=0)
    d.text((408,110),'ATENÇÃO',font=font(78,True),fill=255,anchor='mm')
    d.text((408,255),'PEDIDO COM MAIS',font=font(47,True),fill=0,anchor='mm')
    d.text((408,315),'DE UMA UNIDADE',font=font(47,True),fill=0,anchor='mm')
    size=180
    while font(size,True).getlength(str(total))>710 and size>24:size-=2
    d.text((408,475),str(total),font=font(size,True),fill=0,anchor='mm')
    d.text((408,605),'UNIDADES NO TOTAL',font=font(44,True),fill=0,anchor='mm')
    d.line((50,670,766,670),fill=0,width=4)
    d.text((408,725),'PEDIDO',font=font(30,True),fill=0,anchor='mm')
    y=775
    for line in wrap(row['order_sn'],font(40,True),700):
        d.text((408,y),line,font=font(40,True),fill=0,anchor='mm');y+=48
    d.text((408,965),'CONFERIR TODAS AS UNIDADES',font=font(35,True),fill=0,anchor='mm')
    d.text((408,1020),'ANTES DE FECHAR O PACOTE',font=font(35,True),fill=0,anchor='mm')
    d.text((408,1135),'Acompanha a etiqueta de envio anterior',font=font(25),fill=0,anchor='mm')
    return graphic(im)+'^XA^MMT,Y^PON^MNY^FO0,0^XGR:DEMO.GRF,1,1^FS^PQ1,0,0,N^XZ^XA^IDR:DEMO.GRF^FS^XZ\n',im


def compose(block,row,items):
    m=GRAPHIC.search(block);require(m is not None,'ZPL sem imagem Z64 suportada.')
    raw=zlib.decompress(base64.b64decode(m[4]))
    require(int(m[3])==102 and len(raw)==int(m[2])==124236,'Dimensoes diferentes de 816 x 1218: revisar o layout antes de imprimir.')
    original=ImageOps.invert(Image.frombytes('1',(816,1218),raw).convert('L'))
    footer_empty=original.crop((0,1148,816,1218)).getextrema()==(255,255)
    lines=lines_for(items)
    fits=footer_empty and sum(x[2] for x in lines)<=67
    modified=original.copy();additional=[];previews=[]
    if fits:paint(modified,lines,1151)
    else:
        additional,previews=separation_pages(row,items)
        if footer_empty:
            notice=[('LISTA DE SEPARACAO A SEGUIR',font(23,True),28),('Pedido: '+row['order_sn'],font(21),25)]
            paint(modified,notice,1151)
    require(ImageChops.difference(original,modified).crop((0,0,816,1148)).getbbox() is None,'Dados de envio alterados.')
    shipping=block[:m.start()]+graphic(modified,m[1])+block[m.end():]
    # Keep shipping, quantity warning and all picking pages in one sorting unit.
    if sum(item['qty'] for item in items)>1:
        alert,alert_image=quantity_alert(row,items)
        additional.insert(0,alert);previews.insert(0,alert_image)
    return [shipping]+additional,[modified]+previews

def run(xlsx,archive,destination):
    workbook=openpyxl.load_workbook(xlsx,data_only=True,read_only=True)
    workbook.active.reset_dimensions()
    values=iter(workbook.active.values);headers=next(values)
    require(set(('tracking_number','order_sn','product_info'))<=set(headers),'Colunas obrigatorias ausentes.')
    rows=[dict(zip(headers,r)) for r in values if any(v is not None for v in r)]
    workbook.close()
    require(rows and all(isinstance(r.get('order_sn'),str) and r['order_sn'].strip() and isinstance(r.get('tracking_number'),str) and r['tracking_number'].strip() for r in rows),'Pedido ou rastreio ausente.')
    orders={str(r['tracking_number']):r for r in rows}
    require(len(orders)==len(rows),'Rastreios duplicados na planilha.')
    tracking_aliases={}
    for tracking in orders:
        pickup=re.fullmatch(r'(BR\d{12}[A-Z0-9])SPXLM\d+',tracking)
        printed=pickup[1] if pickup else tracking
        require(printed not in tracking_aliases,'Rastreio impresso ambiguo na planilha.')
        tracking_aliases[printed]=tracking
    with zipfile.ZipFile(archive) as z:
        names=[n for n in z.namelist() if n.lower().endswith(('.txt','.zpl'))]
        require(len(names)==1,'O ZIP deve conter um unico arquivo TXT/ZPL.')
        source=z.read(names[0]).decode('utf-8-sig')
    starts=[m.start() for m in GRAPHIC.finditer(source)]
    require(starts and not source[:starts[0]].strip(),'Inicio do ZPL nao reconhecido.')
    blocks=[source[a:b] for a,b in zip(starts,starts[1:]+[len(source)])]
    bundles=[];seen=set()
    for block in blocks:
        require(block.count('^PQ')==1 and block.count('^XZ')==2,'Estrutura de etiqueta inesperada.')
        m=GRAPHIC.search(block);raw=zlib.decompress(base64.b64decode(m[4]))
        require(int(m[3])==102 and len(raw)==124236,'Layout de origem nao suportado.')
        im=ImageOps.invert(Image.frombytes('1',(816,1218),raw).convert('L'))
        found=set(filter(None,(barcode(im,y) for y in (710,730,750,770)))) & tracking_aliases.keys()
        require(len(found)==1,'Etiqueta sem correspondencia unica na planilha. Confira se XLSX e ZIP pertencem ao mesmo lote; os numeros dos nomes dos arquivos nao precisam ser iguais.')
        tracking=tracking_aliases[found.pop()];require(tracking not in seen,'Etiqueta duplicada.');seen.add(tracking)
        row=orders[tracking];items=items_from(row['product_info'])
        pages,previews=compose(block,row,items)
        key=tuple(sorted(i['sku'].casefold() for i in items))
        bundles.append((key,row,pages))
    missing=set(orders)-seen
    require(not missing,'Faltam etiquetas para '+str(len(missing))+' pedido(s) da planilha. Rastreios: '+', '.join(sorted(missing)[:5]))
    bundles.sort(key=lambda b:b[0])
    result=''.join(page for _,_,pages in bundles for page in pages)
    require(result.count('^PQ')==sum(len(b[2]) for b in bundles),'Contagem de impressoes inconsistente.')
    dest=Path(destination);dest.mkdir(parents=True,exist_ok=True)
    order_count=len({row['order_sn'] for _,row,_ in bundles})
    issued_at=datetime.now().astimezone().strftime('%Y-%m-%d %H-%M-%S')
    filename=f'SHOPEE- {order_count} - {issued_at} - etiquetas-com-separacao'
    path=dest/(filename+'.zpl');path.write_text(result,encoding='ascii')
    with zipfile.ZipFile(dest/(filename+'.zip'),'w',zipfile.ZIP_DEFLATED) as z:z.write(path,path.name)
    with (dest/'ordem-de-impressao.csv').open('w',encoding='utf-8-sig',newline='') as f:
        writer=csv.writer(f,delimiter=';');writer.writerow(['Pedido','Rastreio','Primeira etiqueta','Ultima etiqueta','Etiquetas de separacao','Etiquetas de alerta'])
        position=1
        for _,row,pages in bundles:
            alert=int(sum(i['qty'] for i in items_from(row['product_info']))>1)
            writer.writerow([row['order_sn'],row['tracking_number'],position,position+len(pages)-1,len(pages)-1-alert,alert]);position+=len(pages)
    alerts=sum(sum(i['qty'] for i in items_from(row['product_info']))>1 for _,row,_ in bundles)
    return dict(pedidos=order_count,etiquetas=sum(len(b[2]) for b in bundles),separacao=sum(len(b[2])-1 for b in bundles)-alerts,alertas=alerts,arquivo=str(path))

if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('planilha');parser.add_argument('zip_original');parser.add_argument('pasta_saida')
    args=parser.parse_args()
    print(run(args.planilha,args.zip_original,args.pasta_saida))

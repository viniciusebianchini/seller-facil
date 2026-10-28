"""Aplicativo local: nenhum dado de pedidos sai deste computador."""
import argparse
import io
import json
import mimetypes
import secrets
import shutil
import sys
import threading
import webbrowser
import zipfile
from email.parser import BytesParser
from email.policy import default
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlsplit, unquote

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / 'processor'))
import openpyxl
import processar_lote

LIMIT = 80 * 1024 * 1024
LOCK = threading.Lock()
TOKEN = secrets.token_urlsafe(32)


def inspect_zip(data):
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        if sum(f.file_size for f in archive.infolist()) > 180 * 1024 * 1024:
            raise ValueError('Arquivo compactado excede o limite de 180 MB descompactados.')


def process_files(files):
    files = [(n.strip(), b) for n, b in files]
    sheets = [(n, b) for n, b in files if n.lower().endswith('.xlsx')]
    archives = [(n, b) for n, b in files if n.lower().endswith('.zip')]
    if not sheets or not archives:
        raise ValueError('Inclua pelo menos uma planilha XLSX e um ZIP original da Shopee.')
    if len(sheets) + len(archives) != len(files):
        raise ValueError('Selecione apenas arquivos XLSX e ZIP.')
    job = secrets.token_hex(12)
    folder = ROOT / 'dados' / job
    folder.mkdir(parents=True)
    try:
        book = openpyxl.Workbook()
        table = book.active
        columns = ['tracking_number', 'order_sn', 'product_info']
        table.append(columns)
        seen = set()
        for name, data in sheets:
            inspect_zip(data)
            source = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
            try:
                source.active.reset_dimensions()
                rows = iter(source.active.values)
                headers = tuple(v.strip().lstrip('\ufeff') if isinstance(v, str) else v for v in next(rows, ()))
                if not set(columns) <= set(headers):
                    raise ValueError(f'{name}: não é uma lista de empacotamento compatível.')
                for values in rows:
                    if not any(v is not None for v in values):
                        continue
                    row = dict(zip(headers, values))
                    for key in ('order_sn', 'tracking_number'):
                        if isinstance(row.get(key), str):
                            row[key] = row[key].strip()
                    tracking = row['tracking_number']
                    if tracking in seen:
                        raise ValueError(f'Rastreio duplicado nas planilhas: {tracking}. Remova o arquivo repetido.')
                    seen.add(tracking)
                    try:
                        processar_lote.etiquetas.items_from(row['product_info'])
                    except ValueError as error:
                        raise ValueError(f"Pedido {row['order_sn']}: {error}") from error
                    table.append([row.get(key) for key in columns])
            finally:
                source.close()
        xlsx = folder / 'listas-do-lote.xlsx'
        book.save(xlsx)
        chunks = []
        for name, data in archives:
            inspect_zip(data)
            with zipfile.ZipFile(io.BytesIO(data)) as archive:
                names = [n for n in archive.namelist() if n.lower().endswith(('.zpl', '.txt')) and not n.startswith('__MACOSX/') and not Path(n).name.startswith('._')]
                label_chunks = []
                for member in names:
                    try:
                        content = archive.read(member).decode('utf-8-sig').strip()
                    except UnicodeDecodeError as error:
                        raise ValueError(f'{name}: {member} não contém texto ZPL compatível.') from error
                    if processar_lote.etiquetas.GRAPHIC.search(content):
                        label_chunks.append(content + '\n')
                    elif member.lower().endswith('.zpl') or '^XA' in content:
                        raise ValueError(f'{name}: layout de etiqueta não suportado em {member}.')
                if not label_chunks:
                    raise ValueError(f'{name}: nenhuma etiqueta ZPL original compatível encontrada no ZIP.')
                chunks.extend(label_chunks)
        merged = folder / 'originais.zip'
        with zipfile.ZipFile(merged, 'w', zipfile.ZIP_DEFLATED) as archive:
            archive.writestr('originais.zpl', ''.join(chunks))
        result = processar_lote.process(xlsx, merged, folder / 'resultados', folder / 'previas')
        xlsx.unlink()
        merged.unlink()
        downloads = {}
        for key in ('pdf', 'word', 'arquivo', 'zip'):
            path = Path(result[key])
            downloads[key] = f'/dados/{job}/resultados/{path.name}'
        package = folder / 'resultados' / 'todos-os-resultados.zip'
        with zipfile.ZipFile(package, 'w', zipfile.ZIP_DEFLATED) as archive:
            for path in package.parent.iterdir():
                if path != package and path.name != 'resumo.json':
                    archive.write(path, path.name)
        downloads['todos'] = f'/dados/{job}/resultados/{package.name}'
        return {key: result[key] for key in ('pedidos', 'etiquetas', 'separacao', 'alertas', 'skus', 'unidades')} | {
            'downloads': downloads,
            'previews': [f'/dados/{job}/previas/etiqueta-{n:03}.png' for n in range(1, result['etiquetas'] + 1)],
        }
    except Exception:
        shutil.rmtree(folder)
        raise


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *args):
        pass

    def send(self, status, data, content_type='application/json; charset=utf-8', attachment=None):
        self.send_response(status)
        self.send_header('Content-Type', content_type)
        self.send_header('Content-Length', str(len(data)))
        self.send_header('Cache-Control', 'no-store')
        self.send_header('X-Content-Type-Options', 'nosniff')
        self.send_header('Content-Security-Policy', "default-src 'self'; img-src 'self' data:; style-src 'self'; script-src 'self'; frame-ancestors 'none'")
        if attachment:
            from urllib.parse import quote
            self.send_header('Content-Disposition', "attachment; filename*=UTF-8''" + quote(attachment))
        self.end_headers()
        self.wfile.write(data)

    def valid_host(self):
        return self.headers.get('Host') in (f'127.0.0.1:{self.server.server_port}', f'localhost:{self.server.server_port}')

    def do_GET(self):
        if not self.valid_host():
            return self.send(403, b'{}')
        path = unquote(urlsplit(self.path).path)
        if path == '/api/session':
            return self.send(200, json.dumps({'token': TOKEN}).encode())
        base = ROOT / ('dados' if path.startswith('/dados/') else 'static')
        relative = path[7:] if path.startswith('/dados/') else path.lstrip('/') or 'index.html'
        target = (base / relative).resolve()
        if not target.is_relative_to(base.resolve()) or not target.is_file():
            return self.send(404, b'{}')
        if base.name == 'dados' and target.suffix not in ('.png', '.pdf', '.docx', '.zpl', '.zip', '.csv'):
            return self.send(404, b'{}')
        self.send(200, target.read_bytes(), mimetypes.guess_type(target.name)[0] or 'application/octet-stream',
                  target.name if base.name == 'dados' and target.suffix != '.png' else None)

    def do_POST(self):
        if not self.valid_host() or self.headers.get('X-App-Token') != TOKEN:
            return self.send(403, b'{}')
        if self.path != '/api/process':
            return self.send(404, b'{}')
        if not LOCK.acquire(blocking=False):
            return self.send(409, json.dumps({'error': 'Outro lote está sendo processado. Aguarde.'}).encode())
        try:
            length = int(self.headers.get('Content-Length', '0'))
            if not 0 < length <= LIMIT:
                raise ValueError('O limite total é 80 MB por envio.')
            content_type = self.headers.get('Content-Type', '')
            if not content_type.startswith('multipart/form-data;'):
                raise ValueError('Formato de envio inválido.')
            self.connection.settimeout(120)
            message = BytesParser(policy=default).parsebytes(
                ('Content-Type: ' + content_type + '\r\nMIME-Version: 1.0\r\n\r\n').encode() + self.rfile.read(length))
            files = [(part.get_filename(), part.get_payload(decode=True)) for part in message.iter_parts() if part.get_filename()]
            if len(files) > 40:
                raise ValueError('Envie no máximo 40 arquivos por lote.')
            result = process_files(files)
            self.send(200, json.dumps(result, ensure_ascii=False).encode())
        except Exception as error:
            message = str(error) if isinstance(error, ValueError) else 'Não foi possível ler os arquivos. Confira se são os originais da Shopee e tente novamente.'
            self.send(400, json.dumps({'error': message}, ensure_ascii=False).encode())
        finally:
            LOCK.release()


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--port', type=int, default=8765)
    parser.add_argument('--no-browser', action='store_true')
    args = parser.parse_args()
    server = ThreadingHTTPServer(('127.0.0.1', args.port), Handler)
    url = f'http://127.0.0.1:{server.server_port}'
    print('Etiquetas Shopee: ' + url, flush=True)
    if not args.no_browser:
        webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        server.server_close()

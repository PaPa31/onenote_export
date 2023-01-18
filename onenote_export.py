import os
import random
import re
import string
import time
import uuid
from fnmatch import fnmatch
from html.parser import HTMLParser
from pathlib import Path
from xml.etree import ElementTree
import subprocess

from datetime import datetime

import click
import flask
import msal
import yaml
from pathvalidate import sanitize_filename
from pathvalidate import sanitize_filepath
from requests_oauthlib import OAuth2Session

graph_url = 'https://graph.microsoft.com/v1.0'
authority_url = 'https://login.microsoftonline.com/common'
scopes = ['Notes.Read', 'Notes.Read.All']
redirect_uri = 'http://localhost:5000/getToken'

app = flask.Flask(__name__)
app.debug = True
app.secret_key = os.urandom(16)

with open('config.yaml') as f:
    config = yaml.safe_load(f)

application = msal.ConfidentialClientApplication(
    config['client_id'],
    authority=authority_url,
    client_credential=config['secret']
)


@app.route("/")
def main():
    resp = flask.Response(status=307)
    resp.headers['location'] = '/login'
    return resp


@app.route("/login")
def login():
    auth_state = str(uuid.uuid4())
    flask.session['state'] = auth_state
    authorization_url = application.get_authorization_request_url(scopes, state=auth_state,
                                                                  redirect_uri=redirect_uri)
    resp = flask.Response(status=307)
    resp.headers['location'] = authorization_url
    return resp


def get_json(graph_client, url, params=None, indent=0):
    values = []
    next_page = url
    while next_page:
        resp = get(graph_client, next_page, params=params, indent=indent).json()
        if 'value' not in resp:
            raise RuntimeError(f'Invalid server response: {resp}')
        values += resp['value']
        next_page = resp.get('@odata.nextLink')
    return values


def get(graph_client, url, params=None, indent=0):
    while True:
        resp = graph_client.get(url, params=params)
        if resp.status_code == 429:
            # We are being throttled due to too many requests.
            # See https://docs.microsoft.com/en-us/graph/throttling
            indent_print(indent, 'Too many requests, waiting 300s and trying again.')
            time.sleep(300)
        elif resp.status_code == 500:
            # In my case, one specific note page consistently gave this status
            # code when trying to get the content. The error was "19999:
            # Something failed, the API cannot share any more information
            # at the time of the request."
            indent_print(indent, 'Error 500, skipping this page.')
            return None
        elif resp.status_code == 504:
            indent_print(indent, 'Request timed out, probably due to a large attachment. Skipping.')
            return None
        else:
            resp.raise_for_status()
            return resp


def download_attachments(graph_client, content, out_dir, page_title, indent=0):
    dir_name = page_title + '.FILES'
    attachment_dir = out_dir / dir_name

    class MyHTMLParser(HTMLParser):
        def handle_starttag(self, tag, attrs):
            self.attrs = {k: v for k, v in attrs}

    def generate_html(tag, props):
        element = ElementTree.Element(tag, attrib=props)
        return ElementTree.tostring(element, encoding='unicode')

    def download_image(tag_match):
        # <img width="843" height="218.5" src="..." data-src-type="image/png" data-fullres-src="..."
        # data-fullres-src-type="image/png" />
        parser = MyHTMLParser()
        parser.feed(tag_match[0])
        props = parser.attrs
        image_url = props.get('data-fullres-src', props['src'])
        image_type = props.get('data-fullres-src-type', props['data-src-type']).split("/")[-1]
        file_name = re.sub(r'.*\-(.+)?\!1-(.+?)\/\$value', r'\1', image_url) + '.' + image_type
        if (attachment_dir / file_name).exists():
            indent_print(indent, f'Image {file_name} already downloaded; skipping.')
        else:
            req = get(graph_client, image_url, indent=indent)
            if req is None:
                return tag_match[0]
            img = req.content
            indent_print(indent, f'Downloaded image of {len(img)} bytes.')
            attachment_dir.mkdir(exist_ok=True)
            with open(attachment_dir / file_name, "wb") as f:
                f.write(img)
        props['src'] = dir_name + "/" + file_name
        props = {k: v for k, v in props.items() if 'data-fullres-src' not in k}
        return generate_html('img', props)

    def download_attachment(tag_match):
        # <object data-attachment="Trig_Cheat_Sheet.pdf" type="application/pdf" data="..."
        # style="position:absolute;left:528px;top:139px" />
        parser = MyHTMLParser()
        parser.feed(tag_match[0])
        props = parser.attrs
        data_url = props['data']
        file_name = props['data-attachment']
        if (attachment_dir / file_name).exists():
            indent_print(indent, f'Attachment {file_name} already downloaded; skipping.')
        else:
            req = get(graph_client, data_url, indent=indent)
            if req is None:
                return tag_match[0]
            data = req.content
            indent_print(indent, f'Downloaded attachment {file_name} of {len(data)} bytes.')
            attachment_dir.mkdir(exist_ok=True)
            with open(attachment_dir / file_name, "wb") as f:
                f.write(data)
        #props['data'] = page_title + "/" + file_name
        props['data'] = dir_name + "/" + file_name
        return generate_html('object', props)

    content = re.sub(r"<img .*?\/>", download_image, content, flags=re.DOTALL)
    content = re.sub(r"<object .*?\/>", download_attachment, content, flags=re.DOTALL)
    content = re.sub(r'<object (.+?)\/>', r'<object \1></object>', content)
    return content


def indent_print(depth, text):
    print('  ' * depth + text)


def filter_items(items, select, name='items', indent=0):
    if not select:
        return items, select
    items = [item for item in items
             if fnmatch(item.get('displayName', item.get('title')).lower(), select[0].lower())]
    if not items:
        indent_print(indent, f'No {name} found matching {select[0]}')
    return items, select[1:]


def download_notebooks(graph_client, path, select=None, indent=0):
    notebooks = get_json(graph_client, f'{graph_url}/me/onenote/notebooks')
    indent_print(0, f'Got {len(notebooks)} notebooks:')
    print(f'---------------------------------------------')
    for n in notebooks:
        print(n.get("displayName"))
    print(f'---------------------------------------------')
    notebooks, select = filter_items(notebooks, select, 'notebooks', indent)
    for nb in notebooks:
        nb_name = nb["displayName"]
        indent_print(indent, f'Opening notebook {nb_name}')
        sections = get_json(graph_client, nb['sectionsUrl'])
        section_groups = get_json(graph_client, nb['sectionGroupsUrl'])
        indent_print(indent + 1, f'Got {len(sections)} sections and {len(section_groups)} section groups.')
        download_sections(graph_client, sections, path / (nb_name + '.SECTIONS'), select, indent=indent + 1)
        download_section_groups(graph_client, section_groups, path / (nb_name + '.SECTIONS'), select, indent=indent + 1)


def download_section_groups(graph_client, section_groups, path, select=None, indent=0):
    section_groups, select = filter_items(section_groups, select, 'section groups', indent)
    for sg in section_groups:
        sg_name = sg["displayName"]
        indent_print(indent, f'Opening section group {sg_name}')
        sections = get_json(graph_client, sg['sectionsUrl'])
        indent_print(indent + 1, f'Got {len(sections)} sections.')
        download_sections(graph_client, sections, path / (sg_name + '.SECTIONS'), select, indent=indent + 1)


def download_sections(graph_client, sections, path, select=None, indent=0):
    sections, select = filter_items(sections, select, 'sections', indent)
    for sec in sections:
        sec_name = sec["displayName"]
        indent_print(indent, f'Opening subsection: {path}/{sec_name}.NOTES')
        pages = get_json(graph_client, sec['pagesUrl'] + '?pagelevel=true')
        indent_print(indent + 1, f'Got {len(pages)} pages.')
        download_pages(graph_client, pages, path / (sec_name + '.NOTES'), select, indent=indent + 1)


def download_pages(graph_client, pages, path, select=None, indent=0):
    pages, select = filter_items(pages, select, 'pages', indent)
    pages = sorted([(page['order'], page) for page in pages])
    level_dirs = [None] * 4
    level_dirs[0] = path
    level_dirs[1] = ''
    level_dirs[2] = ''
    for order, page in pages:
        level = page['level']
        page_title = sanitize_filename(f'{page["title"]}')
        page_title = sanitize_filepath(f'{page_title}')
        indent_print(indent, f'Opening page: {page_title}' + '.html')
        if level == 0:
            page_dir = level_dirs[0]
            level_dirs[1] = page_title + '.NOTES'
        if level == 1:
            page_dir = level_dirs[0] / level_dirs[1]
            level_dirs[2] = page_title + '.NOTES'
        if level == 2:
            page_dir = level_dirs[0] / level_dirs[1] / level_dirs[2]
        download_page(graph_client, page['contentUrl'], page_dir, page_title, indent=indent + 1)

def download_page(graph_client, page_url, path, page_title, indent=0):
    out_html = path / (page_title +'.html')
    if out_html.exists():
        indent_print(indent, 'HTML file already exists; skipping this page')
        return
    path.mkdir(parents=True, exist_ok=True)
    response = get(graph_client, page_url, indent=indent)
    if response is not None:
        content = response.text
        indent_print(indent, f'length:{len(content)}')
        content = re.sub(r'(?: {2})', '&nbsp;&nbsp;', content)
        content = re.sub(r'<div .*?>', '<div style="margin:20px;max-width:624px">', content)
        content = re.sub(r'<body .*?>', '<body style="font-family:Calibri;font-size:14pt;background:#1a1a1a;color:#ddd">', content)
        content = re.sub(r'<p .*?>', '<p>', content)
        content = re.sub(r'<a ', '<a style="color:#abffb4" ', content)
        content = re.sub(r'<iframe (.+?)\/>', r'<iframe \1></iframe>', content)
        content = re.sub(r'<table (.*?)border:1px solid;(.*?)>', r'<table \1\2>', content)
        content = re.sub(r'background-color:[^\"|^;]+[;]*', r'',content)
        #content = re.sub(r'border:1px solid;|border:1px solid\"', r'border:1px solid #3e3e3e;padding:5px',content)
        content = content.replace('￼', f'<span><span>&nbsp;</span><br /></span>')
        content = content.replace(' style="color:#333333"', f'')
        content = content.replace('color:#330099', f'color:#c5a6ff')
        content = content.replace('color:#8e0012', f'color:#ff8f9d')
        content = content.replace('color:#cc3300', f'color:#ffa080')
        content = content.replace('color:black', f'color:#eee')
        content = content.replace('color:#333333', f'color:#ff8860')
        content = content.replace('color:#000088', f'color:#99f')
        content = content.replace('color:#006699', f'color:#39bdff')
        content = content.replace('color:#35586c', f'color:#77ceff')
        content = content.replace('color:#555555', f'color:#b9b9b9')
        content = content.replace('color:#336666', f'color:#7bffff')
        content = content.replace('color:#aa0000', f'color:#99f')
        content = content.replace('color:#dbe5f1', f'color:#313131')
        content = content.replace('color:silver', f'color:#313131')
        content = content.replace('color:#d99694', f'color:#404040')
        content = content.replace('color:#1e4e79', f'color:#37a1ff')
        #content = content.replace('color:#292929', f'color:#868686')
        #content = content.replace('color:#212529', f'color:#737373')
        #content = content.replace('color:#242729', f'color:#737373')
        #content = content.replace('color:#24292e', f'color:#737373')
        content = content.replace('color:#0b0080', f'color:#bfb9ff')
        content = content.replace('</head>', '\t<style>\n\t\t\thtml{width:100%}\n\t\t\tbody>div p {margin-top:0pt;margin-bottom:0pt}\n\t\t\timg{display:block}\n\t\t</style>\n\t</head>')
        content = content.replace('<td style="border:1px solid">', f'<td style="border:1px solid #686868;padding:5px;">')
        hexList = re.findall(r'color:\#[0-9a-f]{6}', content)
        for item in hexList:
            content = re.sub(f"{item}", item.replace('2','c'), content)
        content = re.sub(r'<td style="(.*?)border:1px solid(.*?)">', r'<td style="\1border:1px solid #3e3e3e;padding:5px\2">', content)
        dateRegex = re.compile(r'\d{4}-\d{2}-\d{2}')
        if dateRegex.search(content):
            ddate = dateRegex.search(content)
            date_object = datetime.strptime(ddate.group(0), '%Y-%m-%d').date()
        else:
            date_object = '0000-00-00'
        content = re.sub(r'(<div style="margin:20px;max-width:624px.*?>)', r'\1\n\t\t<time style="color:#595959;text-align:right;display:block;font-style:italic;margin-bottom:40px">%s</time>\n'%date_object, content, 1)
        titleRegex = re.compile(r'<title>(.*?)</title>')
        ttitle = titleRegex.search(content)
        if ttitle:
            #title_object = ttitle.group(1)
            title_object = page_title
            content  = re.sub(r'<title>.*?</title>', r'<title>%s</title>'%title_object, content, 1)
        else:
            title_object = page_title
            title_new = '<title>' + title_object + '</title>'
            head_default = '\n\t\t' + title_new + '\n\t\t<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />\n\t\t<meta name="created" content="0000-00-00T00:00:00.0000000" />\n\t\t<style>\n\t\t\thtml{width:100%}\n\t\t\tbody>div p {margin-top:0pt;margin-bottom:0pt}\n\t\t\timg{display:block}\n\t\t</style>\n\t'
            content = re.sub(r'<html lang="en-US">', r'<html lang="en-US">\n\t<head>%s</head>'%head_default, content)
        content = re.sub(r'(<div style="margin:20px;max-width:624px.*?>)', r'\1\n\t\t<h1 style="margin-bottom: 0pt;font-weight: normal;border-bottom: 1px solid #777;">%s</h1>\n'%title_object, content, 1)
        content = download_attachments(graph_client, content, path, page_title, indent=indent)
        with open(out_html, "w", encoding='utf-8') as f:
            f.write(content)


@app.route("/getToken")
def main_logic():
    code = flask.request.args['code']
    token = application.acquire_token_by_authorization_code(code, scopes=scopes,
                                                            redirect_uri=redirect_uri)
    graph_client = OAuth2Session(token=token)
    download_notebooks(graph_client, app.config['output_path'], app.config['select_path'], indent=0)
    print("Done!")
    return flask.render_template_string('<html>'
                                        '<head><title>Done</title></head>'
                                        '<body><p1><b>Done</b></p1></body>'
                                        '</html>')


@click.command()
@click.option('-s', '--select', default='',
              help='Only convert a subset of notes, given as a slash-separated path. For example '
                   '`-p mynotebook` or `-p mynotebook/mysection/mynote`. Wildcards are supported: '
                   '`-p mynotebook/*/mynote`.')
@click.option('-o', '--outdir', default='output', help='Path to output directory.')
def main_command(select, outdir):
    app.config['select_path'] = [x for x in select.split('/') if x]
    app.config['output_path'] = Path(outdir)
    app.run()


if __name__ == "__main__":
    main_command()

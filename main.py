from fasthtml.common import *
from html2text import HTML2Text

app,rt = fast_app(hdrs=(ScriptX('wordpaste.js'),))

def get_cts(s=''):
    return Div(
        Textarea(s, id="pasteArea", placeholder="Paste HTML here",
                 hx_post="/convert", hx_trigger="paste",
                 hx_vals="js:{'html': transformPastedHTML(event.clipboardData.getData('text/html'))}",
                 hx_target="#main",
                 style="width: 100%; height: 80vh;"))

@rt("/")
async def get():
    return Titled(
        "Word to Markdown Converter",
        Div(P(),get_cts(), id='main'))

@rt("/convert")
async def post(html: str):
    h2t = HTML2Text(bodywidth=5000)
    h2t.ignore_links,h2t.mark_code,h2t.ignore_images = True,True,True
    return Div(
        Button("Copy", id="copyBtn", onclick="copyToClipboard()"),
        get_cts(h2t.handle(html))
    )

serve()


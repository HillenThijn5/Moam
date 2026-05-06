from pathlib import Path
from xml.sax.saxutils import escape as xml_escape
from docxtpl import DocxTemplate, RichText
from PCMail.config.paths import WORD_TEMPLATE_PATH, TEMP_DIR


BROCHURE_FONT = "Proxima Nova A Cond"
BROCHURE_SIZE = 16  # docx half-points; 16 == 8pt


def _apply_hyperlinks(doc: DocxTemplate, ctx: dict) -> dict:
    """
    Convert tuple values into clickable hyperlinks.

    Supported tuple formats:
    - (text, url)
    - (text, url, trailing_text)
    """
    out = {}

    for key, value in ctx.items():
        if (
            isinstance(value, tuple)
            and len(value) in (2, 3)
            and value[0]
            and value[1]
        ):
            text, url = value[0], value[1]
            trailing_text = value[2] if len(value) == 3 else ""
            use_brochure_style = key == "brochure"
            rt = RichText()
            rt.add(
                str(text),
                url_id=doc.build_url_id(str(url)),
                color="0000FF",  # blue
                underline=True,  # underline
                font=BROCHURE_FONT if use_brochure_style else None,
                size=BROCHURE_SIZE if use_brochure_style else None,
            )
            if trailing_text:
                rt.add(
                    str(trailing_text),
                    font=BROCHURE_FONT if use_brochure_style else None,
                    size=BROCHURE_SIZE if use_brochure_style else None,
                )

            out[key] = rt
        else:
            out[key] = value

    return out


def render_word(ctx: dict, series: str) -> str:
    doc = DocxTemplate(str(WORD_TEMPLATE_PATH))

    ctx = _apply_hyperlinks(doc, ctx)
    # docxtpl doesn't XML-escape plain strings; pre-escape so & < > survive in the .docx
    ctx = {k: xml_escape(v) if isinstance(v, str) else v for k, v in ctx.items()}

    doc.render(ctx)

    out_path = Path(TEMP_DIR) / f"Individual Product Notification - series {series}.docx"
    doc.save(str(out_path))
    return str(out_path)








from typing import Callable, List, Optional
import copy

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.text.paragraph import Paragraph
from docx.table import Table

from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl

from talkingdb.logger.console import logger
from talkingdb.models.document.document import DocumentModel
from talkingdb.models.document.layouts.layout import LayoutModel, HeaderModel, FooterModel
from talkingdb.models.document.elements.base.base import RunModel, RunAttributes
from talkingdb.models.document.elements.primitive.paragraph import ParagraphModel, ParagraphStyleModel
from talkingdb.models.document.elements.primitive.table import TableModel, TableCellModel

from .docx_paginate import paginate_docx, PAGINATE_DOCX_ENABLED


class DocxReader:
    """Complete DOCX document reader with structured output and robust error handling"""

    def __init__(self):
        self.doc = None
        self.doc_uid = None
        self.io_buffer = None

    def extract_list_info(self, p: Paragraph):
        pPr = p._p.pPr

        if pPr is None or pPr.numPr is None:
            return False, None, 0

        numPr = pPr.numPr
        level = int(numPr.ilvl.val) if numPr.ilvl is not None else 0

        return True, None, level

    def iter_blocks(self, doc):
        for child in doc.element.body.iterchildren():
            if isinstance(child, CT_P):
                yield Paragraph(child, doc)
            elif isinstance(child, CT_Tbl):
                yield Table(child, doc)

    def merge_runs(self, runs: List[RunModel]) -> List[RunModel]:
        if not runs:
            return []

        merged: List[RunModel] = [runs[0]]

        for run in runs[1:]:
            last = merged[-1]

            if run.attributes == last.attributes:
                # merge text
                last.text += run.text
            else:
                merged.append(run)

        return merged

    def extract_runs(self, p: Paragraph) -> List[RunModel]:
        runs: List[RunModel] = []

        for r in p.runs:
            if not r.text:
                continue

            attr = RunAttributes(
                bold=r.bold,
                italic=r.italic,
                underline=r.underline,
                font_size=r.font.size.pt if r.font.size else None,
                subscript=r.font.subscript,
                superscript=r.font.superscript,
                styles=[]
            )

            if r.style:
                attr.styles.append(r.style.name)
            if r.font.name:
                attr.styles.append(f"font:{r.font.name}")
            if r.font.color.rgb:
                attr.styles.append(f"color:{r.font.color.rgb}")

            runs.append(RunModel(text=r.text, attributes=attr))

        return self.merge_runs(runs)

    def extract_table(self, tbl: Table) -> TableModel:
        model = TableModel()

        for r_idx, row in enumerate(tbl.rows):
            row_model = []

            for cell in row.cells:
                colspan = 1
                rowspan = 1
                colspan = int(cell._tc.grid_span)
                if cell._tc.vMerge == "restart":
                    try:
                        rowspan = cell._tc.bottom
                    except ValueError:
                        logger.warning(
                            f"[{self.doc_uid}] table rowspan fallback: no tc at grid_offset="
                            f"{cell._tc.grid_offset} below row {r_idx}; treating cell as unmerged"
                        )
                        rowspan = 1
                _paragraphs = []
                for p in cell.paragraphs:
                    if p.text.strip():
                        is_list, list_type, level = self.extract_list_info(p)

                        _paragraphs.append(ParagraphModel(
                            style=self.extract_paragraph_style(p),
                            runs=self.extract_runs(p),
                            is_list=is_list,
                            list_type=list_type,
                            list_level=level,
                        ))

                row_model.append(
                    TableCellModel(paragraphs=_paragraphs,
                                   colspan=colspan, rowspan=rowspan)
                )

            model.rows.append(row_model)

        return model

    def extract_paragraph_style(self, p: Paragraph) -> ParagraphStyleModel:
        style = p.style
        if style is not None:
            name = style.name
            font = style.font
        else:
            name = p._p.style or "Normal"
            font = None

        pf = p.paragraph_format

        return ParagraphStyleModel(
            name=name,
            alignment=p.alignment.name if p.alignment else None,
            space_before=pf.space_before.pt if pf.space_before else None,
            space_after=pf.space_after.pt if pf.space_after else None,
            bold=font.bold if font else None,
            italic=font.italic if font else None,
            font_size=font.size.pt if (font and font.size) else None,
        )

    def read_header_footer(self, container) -> List[RunModel]:
        runs = []
        for p in container.paragraphs:
            runs.extend(self.extract_runs(p))
        return runs

    def read_document(
        self,
        io_buffer,
        file_name,
        paginate: bool = True,
        cancel_check: Optional[Callable[[], bool]] = None,
        channel: Optional[str] = None,
        file_hash: Optional[str] = None,
    ) -> DocumentModel:
        self.io_buffer = io_buffer
        self.doc_uid = DocumentModel.make_uid(io_buffer)

        io_buffer.seek(0)
        raw_docx_bytes = io_buffer.read()
        io_buffer.seek(0)

        doc = Document(self.io_buffer)
        model = DocumentModel(filename=file_name)

        section_idx = 0
        section = doc.sections[0]

        def new_layout(sec, prev_layout=None):
            if prev_layout is not None and sec.header.is_linked_to_previous:
                header = copy.deepcopy(prev_layout.header)
            else:
                header = HeaderModel(self.read_header_footer(sec.header))

            if prev_layout is not None and sec.footer.is_linked_to_previous:
                footer = copy.deepcopy(prev_layout.footer)
            else:
                footer = FooterModel(self.read_header_footer(sec.footer))

            return LayoutModel(
                orientation="LANDSCAPE" if sec.orientation == WD_ORIENT.LANDSCAPE else "PORTRAIT",
                header=header,
                footer=footer,
            )

        layout = new_layout(section)
        model.layouts.append(layout)

        for block in self.iter_blocks(doc):
            if isinstance(block, Paragraph):
                pPr = block._p.pPr
                sectPr = pPr.sectPr if pPr is not None else None

                if sectPr is not None:
                    section_idx += 1
                    layout = new_layout(doc.sections[section_idx], prev_layout=layout)
                    model.layouts.append(layout)
                    continue

                is_list, list_type, level = self.extract_list_info(block)

                if block.text.strip():
                    layout.elements.append(
                        ParagraphModel(
                            style=self.extract_paragraph_style(block),
                            runs=self.extract_runs(block),
                            is_list=is_list,
                            list_type=list_type,
                            list_level=level,
                        )
                    )

            elif isinstance(block, Table):
                layout.elements.append(self.extract_table(block))

        model.assign_ids(self.doc_uid)
        model.build_hierarchy()

        if paginate and PAGINATE_DOCX_ENABLED:
            paginate_docx(
                raw_docx_bytes, model, cancel_check=cancel_check,
                channel=channel, file_hash=file_hash,
            )

        return model
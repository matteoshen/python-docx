# encoding: utf-8

"""
|Document| and closely related objects
"""
from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from pandas import DataFrame
from pandas import concat
from numpy import arange
from numpy import isnan
from numpy import nan

from .blkcntnr import BlockItemContainer
from .enum.section import WD_SECTION
from .enum.text import WD_BREAK
from .section import Section, Sections
from .shared import ElementProxy, Emu
from .oxml.table import CT_Tbl
from .oxml.text.paragraph import CT_P
from .table import _Cell, Table
from .text.paragraph import Paragraph
from .utilities import SPECIAL_SEP
from .utilities import SPECIALCHARS_CH
from .utilities import SPECIALCHARS_EN
from .utilities import rmSpecailChar

class Document(ElementProxy):
    """
    WordprocessingML (WML) document. Not intended to be constructed directly.
    Use :func:`docx.Document` to open or create a document.
    """

    __slots__ = ('_part', '__body','_dataframe','_mapping_without_sc',
                 '_block_list','_paragraph_or_table','_block_dataframe_list',
                 '_fulltext','_fulltext_without_sc')

    def __init__(self, element, part):
        super(Document, self).__init__(element)
        self._part = part
        self.__body = None
        self._dataframe = None
        self._mapping_without_sc = None
        self._block_list = []
        self._block_dataframe_list = []
        self._paragraph_or_table = []
        self._fulltext = None
        self._fulltext_without_sc = None

    def add_heading(self, text='', level=1):
        """
        Return a heading paragraph newly added to the end of the document,
        containing *text* and having its paragraph style determined by
        *level*. If *level* is 0, the style is set to `Title`. If *level* is
        1 (or omitted), `Heading 1` is used. Otherwise the style is set to
        `Heading {level}`. Raises |ValueError| if *level* is outside the
        range 0-9.
        """
        if not 0 <= level <= 9:
            raise ValueError("level must be in range 0-9, got %d" % level)
        style = 'Title' if level == 0 else 'Heading %d' % level
        return self.add_paragraph(text, style)

    def add_page_break(self):
        """
        Return a paragraph newly added to the end of the document and
        containing only a page break.
        """
        paragraph = self.add_paragraph()
        paragraph.add_run().add_break(WD_BREAK.PAGE)
        return paragraph

    def add_paragraph(self, text='', style=None):
        """
        Return a paragraph newly added to the end of the document, populated
        with *text* and having paragraph style *style*. *text* can contain
        tab (``\\t``) characters, which are converted to the appropriate XML
        form for a tab. *text* can also include newline (``\\n``) or carriage
        return (``\\r``) characters, each of which is converted to a line
        break.
        """
        return self._body.add_paragraph(text, style)

    def add_picture(self, image_path_or_stream, width=None, height=None):
        """
        Return a new picture shape added in its own paragraph at the end of
        the document. The picture contains the image at
        *image_path_or_stream*, scaled based on *width* and *height*. If
        neither width nor height is specified, the picture appears at its
        native size. If only one is specified, it is used to compute
        a scaling factor that is then applied to the unspecified dimension,
        preserving the aspect ratio of the image. The native size of the
        picture is calculated using the dots-per-inch (dpi) value specified
        in the image file, defaulting to 72 dpi if no value is specified, as
        is often the case.
        """
        run = self.add_paragraph().add_run()
        return run.add_picture(image_path_or_stream, width, height)

    def add_section(self, start_type=WD_SECTION.NEW_PAGE):
        """
        Return a |Section| object representing a new section added at the end
        of the document. The optional *start_type* argument must be a member
        of the :ref:`WdSectionStart` enumeration, and defaults to
        ``WD_SECTION.NEW_PAGE`` if not provided.
        """
        new_sectPr = self._element.body.add_section_break()
        new_sectPr.start_type = start_type
        return Section(new_sectPr)

    def add_table(self, rows, cols, style=None):
        """
        Add a table having row and column counts of *rows* and *cols*
        respectively and table style of *style*. *style* may be a paragraph
        style object or a paragraph style name. If *style* is |None|, the
        table inherits the default table style of the document.
        """
        table = self._body.add_table(rows, cols, self._block_width)
        table.style = style
        return table

    @property
    def dataframe(self):
        """
        return DataFrame containing block level info
        """
        return self._dataframe

    @property
    def blocks(self):
        """
        return list of all blocks(paragraph or table) in the doc
        """
        return self._block_list

    @property
    def block_dataframe_list(self):
        """
        return list of all blocks(paragraph or table) in the doc
        """
        return self._block_dataframe_list

    @property
    def fulltext(self):
        """
        return full text string of doc
        """
        if self._dataframe is None:
            self.parse()
        self._fulltext = self._dataframe['string'].sum()
        return self._fulltext

    @property
    def fulltext_without_sc(self):
        """
        return full text string WITHOUT special characters of doc
        """
        self._removeSpecailChar()
        self._fulltext_without_sc = self._mapping_without_sc['char'].sum()
        return self._fulltext_without_sc

    @property
    def paragraph_or_table(self):
        """
        return list of with elements in {'p','t'} informing
            if the n-th block is paragraph or table
        """
        return self._paragraph_or_table

    @property
    def mapping_without_sc(self):
        """
        return DataFrame with mapping info on removing speical characters:
        Special Characters:
            SPECIALCHARS_EN = r'~!@#$%^&*()_-+={}[]|\`:\"\'<>?/.,;'
            SPECIALCHARS_CH = r'·~！@#￥%……&*（）——+-={}|【】：“‘；：”’《》，。？、'
            SPECIAL_SEP = [' ','\n','\t','\u3000']
        """
        return self._mapping_without_sc


    @property
    def core_properties(self):
        """
        A |CoreProperties| object providing read/write access to the core
        properties of this document.
        """
        return self._part.core_properties

    @property
    def inline_shapes(self):
        """
        An |InlineShapes| object providing access to the inline shapes in
        this document. An inline shape is a graphical object, such as
        a picture, contained in a run of text and behaving like a character
        glyph, being flowed like other text in a paragraph.
        """
        return self._part.inline_shapes

    @property
    def paragraphs(self):
        """
        A list of |Paragraph| instances corresponding to the paragraphs in
        the document, in document order. Note that paragraphs within revision
        marks such as ``<w:ins>`` or ``<w:del>`` do not appear in this list.
        """
        return self._body.paragraphs

    @property
    def part(self):
        """
        The |DocumentPart| object of this document.
        """
        return self._part

    def save(self, path_or_stream):
        """
        Save this document to *path_or_stream*, which can be either a path to
        a filesystem location (a string) or a file-like object.
        """
        self._part.save(path_or_stream)

    @property
    def sections(self):
        """
        A |Sections| object providing access to each section in this
        document.
        """
        return Sections(self._element)

    @property
    def settings(self):
        """
        A |Settings| object providing access to the document-level settings
        for this document.
        """
        return self._part.settings

    @property
    def styles(self):
        """
        A |Styles| object providing access to the styles in this document.
        """
        return self._part.styles

    @property
    def tables(self):
        """
        A list of |Table| instances corresponding to the tables in the
        document, in document order. Note that only tables appearing at the
        top level of the document appear in this list; a table nested inside
        a table cell does not appear. A table within revision marks such as
        ``<w:ins>`` or ``<w:del>`` will also not appear in the list.
        """
        return self._body.tables

    @property
    def _block_width(self):
        """
        Return a |Length| object specifying the width of available "writing"
        space between the margins of the last section of this document.
        """
        section = self.sections[-1]
        return Emu(
            section.page_width - section.left_margin - section.right_margin
        )

    @property
    def _body(self):
        """
        The |_Body| instance containing the content for this document.
        """
        if self.__body is None:
            self.__body = _Body(self._element.body, self)
        return self.__body

    def _iter_block_items(self):
        """
        Yield each paragraph and table child within *parent*, in document order.
        Each returned value is an instance of either Table or Paragraph. *parent*
        would most commonly be a reference to a main Document object, but
        also works for a _Cell object, which itself can contain paragraphs and tables.
        """
        if isinstance(self, Document):
            parent_elm = self.element.body
        elif isinstance(self, _Cell):
            parent_elm = self._tc
        else:
            raise ValueError("something's not right")

        for child in parent_elm.iterchildren():
            if isinstance(child, CT_P):
                yield Paragraph(child, self)
            elif isinstance(child, CT_Tbl):
                yield Table(child, self)

    def iter_block_items(self):
        """
        get self._block_list filled
        """
        for i in self._iter_block_items():
            self._block_list.append(i)
            self._block_dataframe_list.append(DataFrame())
            if isinstance(i,Paragraph):
                self._paragraph_or_table.append(1) # 1 means paragraph
            else:
                self._paragraph_or_table.append(0) # 0 means table

    def _parse_block(self,idx):
        """
        Parse self._block_list[idx] into pandas.DataFrame.
        """
        block_tmp = self._block_list[idx]
        blocktype = self._paragraph_or_table[idx]
        paragraph_count = sum(self._paragraph_or_table[:idx+1])
        table_count = idx + 1 - paragraph_count
        df = DataFrame()
        # paragraph
        if blocktype==1:
            l_runText = [r.text for r in block_tmp.runs]
            l_runID = arange(len(l_runText))
            df = DataFrame({'string':l_runText,
                                'run_ID':l_runID},index=l_runID)
            df['paragraph_ID'] = paragraph_count - 1 # 0-starting index 
        # table
        if blocktype==0:
            row_count = 0
            for row in block_tmp.rows:
                cell_count = 0
                for cell in row.cells:
                    cell_para_count = 0
                    for p in cell.paragraphs:
                        l_runText = [r.text for r in p.runs]
                        l_runID = arange(len(l_runText))            
                        df = DataFrame({'string':l_runText,
                                            'run_ID':l_runID},index=l_runID)
                        df['table_ID'] = table_count - 1 # 0-starting index
                        df['row_ID'] = row_count
                        df['cell_ID'] = cell_count
                        df['paragraph_ID'] = cell_para_count 
                        cell_para_count += 1
                    cell_count += 1
                row_count += 1
        df['block_ID'] = idx
        self._block_dataframe_list[idx] = df

    def parse(self):
        """
        Parse a Document into pandas.DataFrame.
        if the document does NOT contain any tables, the return DataFrame's columns will be:
            ['string','block_ID','paragraph_ID','run_ID']
            each ID starts with 0
        if the document DOES contain tables, in this case the return DataFrame's columns will be:
            ['string','block_ID','table_ID','row_ID',
             'cell_ID','paragraph_ID','run_ID']
            each ID starts with 0
        """
        self.iter_block_items()
        for idx in arange(len(self._block_list)):
            self._parse_block(idx)
        self._dataframe = concat(self._block_dataframe_list, ignore_index=True)

    def _reparse_block(self,idx):
        self._parse_block(idx)
        self._dataframe = concat(self._block_dataframe_list, ignore_index=True)

    def _highlight_basic(self,
                        idx,
                        start_pos_relative,
                        end_pos_relative,
                        highlight_color):
        """
        Inner method: HighLight doc with fixed
            DataFrame, idx, start_pos(relative), end_pos(relative) and fixed color
        """
        df = self._dataframe
        # paragraph or table
        blocktype = None
        if 'table_ID' not in df.columns:
            blocktype = 'paragraph'
        elif isnan(df.loc[idx,'table_ID']):
            blocktype = 'paragraph'
        else:        
            blocktype = 'table'
            
        # paragraph
        if blocktype=='paragraph':
            p = self.paragraphs[df.loc[idx,'paragraph_ID']]
            r = p.runs[df.loc[idx,'run_ID']]
            head = r.text[:start_pos_relative]
            mid = r.text[start_pos_relative:end_pos_relative]
            tail = r.text[end_pos_relative:]
            # head
            if head>'':
                r.insert_run_before(text= head,
                                    style=r.style,
                                    font_from_run=r)
            # mid
            r.insert_run_before(text=mid,
                                style=r.style,
                                highlight_color=highlight_color,                                
                                font_from_run=r)
            # tail
            if tail>'':
                r.insert_run_before(text= tail,
                                    style=r.style,
                                    font_from_run=r)
            # delete run
            r.delete_run()
            # arrange self._dataframe
            self._reparse_block(df.loc[idx,'block_ID'])
        # table
        if blocktype=='table':
            table = self.tables[int(df.loc[idx,'table_ID'])]
            row = table.rows[int(df.loc[idx,'row_ID'])]
            cell = row.cells[int(df.loc[idx,'cell_ID'])]
            p = cell.paragraphs[int(df.loc[idx,'paragraph_ID'])]
            r = p.runs[int(df.loc[idx,'run_ID'])]
            head = r.text[:start_pos_relative]
            mid = r.text[start_pos_relative:end_pos_relative]
            tail = r.text[end_pos_relative:]
            # head
            if head>'':
                r.insert_run_before(text= head,
                                    style=r.style,
                                    font_from_run=r)
            # mid
            r.insert_run_before(text=mid,
                                style=r.style,
                                highlight_color=highlight_color,                                
                                font_from_run=r)
            # tail
            if tail>'':
                r.insert_run_before(text= tail,
                                    style=r.style,
                                    font_from_run=r)
            # delete run
            r.delete_run()
            # arrange self._dataframe
            self._reparse_block(df.loc[idx,'block_ID'])

    def highlight(self, position_list, highlight_color, rmSC=False):
        """
        for each (start_pos, end_pos) in @position_list:
        highlight from @start_pos(included,0-starting-index) to 
                        @end_pos with color @highlight_color(included,0-starting-index)

        @rmSC: default=True, 
            if True position_list's postions reflex positions WITHOUT special characters
        
        return Document object
        """
        if self._dataframe is None:
            self.parse()
        # do _removeSpecailChar if needed
        if rmSC:
            self._removeSpecailChar()

        for pos in position_list:
            if nan in pos:
                continue
            if pos[0]>pos[1]:
                raise ValueError('end_pos <%i> should be BIGGER than start_pos <%i>'%(int(pos[1]),int(pos[0])))

            df = self._dataframe.copy()
            df['len_string'] = df['string'].apply(lambda x:len(x))
            df['last_num'] = df['len_string'].cumsum() # last 1-staring num
            df['first_num'] = df['last_num'].shift(1) + 1 # first 1-starting num

            if df.shape[0]>0:
                df.loc[0,'first_num'] = 1

            # 1-starting
            if not rmSC:
                start_num = pos[0] + 1
                end_num = pos[1] + 1
            else:
                start_num = self._mapping_without_sc.loc[pos[0],'index'] + 1
                end_num = self._mapping_without_sc.loc[pos[1],'index'] + 1
            coverd_idx = df[~(start_num > df['last_num']) & 
                            ~(end_num < df['first_num'])].index
                            
            for idx in coverd_idx:
                # 0-starting
                start_pos_relative = max(start_num - df.loc[idx,'first_num'],0)
                end_pos_relative = max(end_num - df.loc[idx,'first_num'],0)
                self._highlight_basic(idx,
                                        int(start_pos_relative),
                                        int(end_pos_relative) + 1,
                                        highlight_color)

    def _removeSpecailChar(self):
        """
        从一个self._dataframe中的指定col里移除特殊字符
        """
        if self._dataframe is None:
            self.parse()
        str_tmp = self._dataframe['string'].sum()
        df_tmp = DataFrame({'char':[c for c in str_tmp]},
                            index=arange(len(str_tmp)))
        df_tmp['char'] = df_tmp['char'].apply(rmSpecailChar)
        df_tmp = df_tmp[df_tmp['char']!='']
        self._mapping_without_sc = df_tmp.reset_index()
                
class _Body(BlockItemContainer):
    """
    Proxy for ``<w:body>`` element in this document, having primarily a
    container role.
    """
    def __init__(self, body_elm, parent):
        super(_Body, self).__init__(body_elm, parent)
        self._body = body_elm

    def clear_content(self):
        """
        Return this |_Body| instance after clearing it of all content.
        Section properties for the main document story, if present, are
        preserved.
        """
        self._body.clear_content()
        return self

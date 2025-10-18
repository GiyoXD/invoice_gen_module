from pydantic import BaseModel, Field
from typing import Optional, Dict

class FontModel(BaseModel):
    name: Optional[str] = None
    size: Optional[float] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    color: Optional[str] = None

class AlignmentModel(BaseModel):
    horizontal: Optional[str] = None
    vertical: Optional[str] = None
    wrap_text: Optional[bool] = Field(default=False, alias='wrapText')

class BorderStyleModel(BaseModel):
    style: Optional[str] = None
    color: Optional[str] = None

class ColumnStyleModel(BaseModel):
    font: Optional[FontModel] = None
    alignment: Optional[AlignmentModel] = None
    number_format: Optional[str] = Field(default=None, alias='numberFormat')

class StylingConfigModel(BaseModel):
    default_font: Optional[FontModel] = Field(default=None, alias='defaultFont')
    default_alignment: Optional[AlignmentModel] = Field(default=None, alias='defaultAlignment')
    header_font: Optional[FontModel] = Field(default=None, alias='headerFont')
    header_alignment: Optional[AlignmentModel] = Field(default=None, alias='headerAlignment')
    column_id_styles: Dict[str, ColumnStyleModel] = Field(default={}, alias='columnIdStyles')
    column_ids_with_full_grid: Optional[list[str]] = Field(default=None, alias='columnIdsWithFullGrid')
    force_text_format_ids: Optional[list[str]] = Field(default=None, alias='forceTextFormatIds')
    column_id_widths: Optional[Dict[str, float]] = Field(default=None, alias='columnIdWidths')
    row_heights: Optional[Dict[str, float]] = Field(default=None, alias='rowHeights')

# powerpoint_service/models.py
from pydantic import BaseModel
from typing import List

class Slide(BaseModel):
    title: str
    content: str

class PresentationRequest(BaseModel):
    title: str
    author: str
    slides: List[Slide]

from docx import Document
from docx.shared import Pt
from typing import List, Dict, Any, Optional


class Section:
    def __init__(self, title: str, level: int, content: str = "", tables: List[Dict] = None):
        self.title = title
        self.level = level
        self.content = content
        self.tables = tables if tables else []
        self.children = []

    def add_child(self, child: 'Section'):
        self.children.append(child)

    def to_dict(self) -> Dict[str, Any]:
        return {
            'title': self.title,
            'level': self.level,
            'content': self.content,
            'tables': self.tables,
            'children': [child.to_dict() for child in self.children]
        }


class WordParser:
    def __init__(self, doc_path: str):
        self.doc_path = doc_path
        self.document = None
        self.sections = []
        self.style_level_map = {
            'Heading 1': 1,
            'Heading 2': 2,
            'Heading 3': 3,
            'Heading 4': 4,
            'Heading 5': 5,
            'Heading 6': 6,
        }

    def load_document(self):
        self.document = Document(self.doc_path)

    def _get_paragraph_level(self, paragraph) -> Optional[int]:
        style_name = paragraph.style.name
        return self.style_level_map.get(style_name)

    def _parse_table(self, table) -> List[List[str]]:
        table_data = []
        for row in table.rows:
            row_data = []
            for cell in row.cells:
                row_data.append(cell.text.strip())
            table_data.append(row_data)
        return table_data

    def extract_sections(self) -> List[Section]:
        if not self.document:
            self.load_document()

        self.sections = []
        stack = []
        current_section = None
        
        all_elements = []
        for paragraph in self.document.paragraphs:
            all_elements.append(('paragraph', paragraph))
        
        for table in self.document.tables:
            all_elements.append(('table', table))
        
        all_elements.sort(key=lambda x: self._get_element_position(x[1]))

        for elem_type, element in all_elements:
            if elem_type == 'paragraph':
                paragraph = element
                level = self._get_paragraph_level(paragraph)
                
                if level is not None:
                    new_section = Section(title=paragraph.text.strip(), level=level)
                    
                    while stack and stack[-1].level >= level:
                        stack.pop()
                    
                    if stack:
                        stack[-1].add_child(new_section)
                    else:
                        self.sections.append(new_section)
                    
                    stack.append(new_section)
                    current_section = new_section
                elif current_section is not None:
                    current_section.content += paragraph.text.strip() + '\n'
            elif elem_type == 'table':
                if current_section is not None:
                    table_data = self._parse_table(element)
                    current_section.tables.append(table_data)

        for section in self.sections:
            section.content = section.content.strip()

        return self.sections

    def _get_element_position(self, element):
        if hasattr(element, '_element'):
            elem = element._element
        else:
            elem = element
        
        for i, child in enumerate(self.document.element.body):
            if child is elem:
                return i
        return float('inf')

    def find_section_by_name(self, name: str, sections: Optional[List[Section]] = None) -> Optional[Section]:
        if sections is None:
            sections = self.sections

        for section in sections:
            if section.title == name:
                return section
            
            if section.children:
                found = self.find_section_by_name(name, section.children)
                if found:
                    return found
        
        return None

    def get_section_content(self, section: Section) -> Dict[str, Any]:
        return {
            'title': section.title,
            'level': section.level,
            'content': section.content,
            'tables': section.tables,
            'children': [self.get_section_content(child) for child in section.children]
        }

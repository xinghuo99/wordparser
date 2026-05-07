import pytest
import os
import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from word_parser import WordParser, Section
from document_creator import create_test_document


TEST_DOC_PATH = 'test_document.docx'


@pytest.fixture(scope='module')
def test_document():
    create_test_document(TEST_DOC_PATH)
    yield TEST_DOC_PATH
    if os.path.exists(TEST_DOC_PATH):
        os.remove(TEST_DOC_PATH)


def test_extract_sections(test_document):
    parser = WordParser(test_document)
    sections = parser.extract_sections()
    
    assert len(sections) == 2
    
    assert sections[0].title == '第一章 概述'
    assert sections[0].level == 1
    assert '第一章的正文内容' in sections[0].content
    assert sections[0].tables == []
    
    assert sections[1].title == '第二章 技术方案'
    assert sections[1].level == 1
    assert '技术方案的设计' in sections[1].content


def test_section_hierarchy(test_document):
    parser = WordParser(test_document)
    sections = parser.extract_sections()
    
    chapter1 = sections[0]
    assert len(chapter1.children) == 2
    
    section1_1 = chapter1.children[0]
    assert section1_1.title == '1.1 项目背景'
    assert section1_1.level == 2
    assert len(section1_1.tables) == 1
    assert section1_1.tables[0][0] == ['序号', '项目名称', '状态']
    
    section1_2 = chapter1.children[1]
    assert section1_2.title == '1.2 项目目标'
    assert section1_2.level == 2
    
    chapter2 = sections[1]
    assert len(chapter2.children) == 2
    
    section2_1 = chapter2.children[0]
    assert section2_1.title == '2.1 技术选型'
    assert len(section2_1.children) == 2
    
    section2_1_1 = section2_1.children[0]
    assert section2_1_1.title == '2.1.1 前端技术'
    assert section2_1_1.level == 3
    assert len(section2_1_1.tables) == 1
    
    section2_1_2 = section2_1.children[1]
    assert section2_1_2.title == '2.1.2 后端技术'
    assert section2_1_2.level == 3


def test_find_section_by_name(test_document):
    parser = WordParser(test_document)
    parser.extract_sections()
    
    found = parser.find_section_by_name('第一章 概述')
    assert found is not None
    assert found.title == '第一章 概述'
    
    found = parser.find_section_by_name('1.1 项目背景')
    assert found is not None
    assert found.title == '1.1 项目背景'
    assert len(found.tables) == 1
    
    found = parser.find_section_by_name('2.1.1 前端技术')
    assert found is not None
    assert found.title == '2.1.1 前端技术'
    
    found = parser.find_section_by_name('不存在的章节')
    assert found is None


def test_find_section_with_children(test_document):
    parser = WordParser(test_document)
    parser.extract_sections()
    
    found = parser.find_section_by_name('2.1 技术选型')
    assert found is not None
    assert len(found.children) == 2
    
    child_titles = [child.title for child in found.children]
    assert '2.1.1 前端技术' in child_titles
    assert '2.1.2 后端技术' in child_titles


def test_get_section_content(test_document):
    parser = WordParser(test_document)
    parser.extract_sections()
    
    section = parser.find_section_by_name('第一章 概述')
    content = parser.get_section_content(section)
    
    assert content['title'] == '第一章 概述'
    assert content['level'] == 1
    assert '正文内容' in content['content']
    assert len(content['children']) == 2
    
    child_titles = [child['title'] for child in content['children']]
    assert '1.1 项目背景' in child_titles
    assert '1.2 项目目标' in child_titles


def test_table_extraction(test_document):
    parser = WordParser(test_document)
    parser.extract_sections()
    
    section = parser.find_section_by_name('1.1 项目背景')
    assert len(section.tables) == 1
    
    table = section.tables[0]
    assert len(table) == 3
    assert table[0] == ['序号', '项目名称', '状态']
    assert table[1] == ['1', '项目A', '进行中']
    assert table[2] == ['2', '项目B', '已完成']

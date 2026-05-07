from src.word_parser import WordParser
from src.document_creator import create_test_document
import os


def print_section_info(section, indent=0):
    """递归打印章节信息"""
    prefix = "  " * indent
    print(f"{prefix}├── [{section.level}] {section.title}")
    
    if section.content:
        content_preview = section.content[:50] + "..." if len(section.content) > 50 else section.content
        print(f"{prefix}│   └── 内容: {content_preview}")
    
    if section.tables:
        print(f"{prefix}│   └── 表格数量: {len(section.tables)}")
        for i, table in enumerate(section.tables):
            print(f"{prefix}│       └── 表格{i+1}: {len(table)}行 x {len(table[0])}列")
    
    for child in section.children:
        print_section_info(child, indent + 1)


def main():
    test_doc = "test_document.docx"
    
    if not os.path.exists(test_doc):
        print(f"创建测试文档: {test_doc}")
        create_test_document(test_doc)
    
    print("\n=== 加载并解析Word文档 ===")
    parser = WordParser(test_doc)
    sections = parser.extract_sections()
    
    print(f"共提取到 {len(sections)} 个顶级章节")
    print("\n=== 章节结构 ===")
    for section in sections:
        print_section_info(section)
    
    print("\n=== 按名称查找章节 ===")
    search_name = "1.1 项目背景"
    found = parser.find_section_by_name(search_name)
    if found:
        print(f"找到章节: [{found.level}] {found.title}")
        print(f"内容预览: {found.content[:100]}...")
        if found.tables:
            print(f"包含表格:")
            for row in found.tables[0]:
                print(f"  {row}")
    else:
        print(f"未找到章节: {search_name}")
    
    print("\n=== 获取章节完整内容（含子章节） ===")
    chapter2 = parser.find_section_by_name("第二章 技术方案")
    if chapter2:
        content = parser.get_section_content(chapter2)
        print(f"章节: {content['title']}")
        print(f"级别: {content['level']}")
        print(f"子章节数量: {len(content['children'])}")
        for child in content['children']:
            print(f"  └── {child['title']}")
    
    print("\n=== 提取页眉内容 ===")
    headers = parser.extract_headers()
    if headers:
        print(f"共提取到 {len(headers)} 个节的页眉")
        for section_info in headers:
            print(f"\n节 {section_info['section_index'] + 1}:")
            for header in section_info['headers']:
                print(f"  类型: {header['type']}")
                print(f"  内容: {header['content']}")
                if header['tables']:
                    print(f"  表格数量: {len(header['tables'])}")
    else:
        print("文档中没有页眉内容")
    
    print("\n=== 获取所有页眉文本 ===")
    all_header_text = parser.get_all_headers_text()
    print(f"所有页眉内容:\n{all_header_text}")
    
    print("\n=== 获取整个文档内容（文字+表格转文字） ===")
    full_text = parser.get_full_document_text(include_headers=True)
    print(full_text)
    
    print("\n=== 按章节名称获取内容（包含子章节） ===")
    section_text = parser.get_section_text_by_name('第一章 概述')
    print(section_text)
    
    print("\n=== 按章节名称获取内容（不包含子章节） ===")
    section_text_no_child = parser.get_section_text_by_name('第一章 概述', include_children=False)
    print(section_text_no_child)
    
    print("\n=== 按章节名称获取子章节内容 ===")
    section_text_child = parser.get_section_text_by_name('1.1 项目背景')
    print(section_text_child)
    
    print("\n=== 获取修订记录内容 ===")
    revision_text = parser.get_revision_records()
    if revision_text:
        print(revision_text)
    else:
        print("未找到修订记录章节")


if __name__ == "__main__":
    main()
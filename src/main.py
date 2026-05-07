from word_parser import WordParser

def main():
    parser = WordParser('test.docx')
    sections = parser.extract_sections()

    # 查找章节
    section = parser.find_section_by_name('第一章 概述')
    content = parser.get_section_content(section)
    print(content)  

if __name__ == '__main__':
    main()
from docx import Document
from docx.oxml.ns import qn
import lxml.etree


def deep_extract_text(element, nsmap=None):
    """
    递归提取一个 OOXML 元素及其所有子元素中的文本。
    按文档顺序拼接，保留段落和表格的基本分隔。
    """
    if nsmap is None:
        nsmap = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
            'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
            'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
        }

    parts = []

    # 处理元素本身的文本（极少情况，但以防万一）
    if element.text:
        parts.append(element.text)

    for child in element:
        # 简化标签名（去掉命名空间前缀）
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

        # ---------- 1. 直接文本节点 ----------
        if tag == 't':                 # w:t
            if child.text:
                parts.append(child.text)
        elif tag == 'instrText':      # w:instrText 域代码
            if child.text:
                parts.append(child.text)
        elif tag == 'delText':        # w:delText 修订删除文本
            if child.text:
                parts.append(child.text)

        # ---------- 2. 换行符 ----------
        elif tag == 'br':             # w:br
            parts.append('\n')

        # ---------- 3. 段落 ----------
        elif tag == 'p':              # w:p
            parts.append(deep_extract_text(child, nsmap))
            parts.append('\n')        # 段尾添加换行

        # ---------- 4. 表格 ----------
        elif tag == 'tbl':            # w:tbl
            parts.append(deep_extract_text(child, nsmap))
        elif tag == 'tr':             # w:tr 行
            parts.append(deep_extract_text(child, nsmap))
            parts.append('\n')        # 一行结束换行
        elif tag == 'tc':             # w:tc 单元格
            parts.append(deep_extract_text(child, nsmap))
            parts.append('\t')        # 单元格间用制表符分隔

        # ---------- 5. 文本框 ----------
        elif tag == 'txbxContent':    # w:txbxContent
            parts.append(deep_extract_text(child, nsmap))

        # ---------- 6. 结构化文档标签 ----------
        elif tag == 'sdt':            # w:sdt
            # 直接递归，sdt 中通常是 w:sdtContent 包裹的内容
            parts.append(deep_extract_text(child, nsmap))

        # ---------- 7. 兼容性备用内容 ----------
        elif tag == 'AlternateContent':  # mc:AlternateContent
            fallback = child.find(qn('mc:Fallback'))
            if fallback is not None:
                parts.append(deep_extract_text(fallback, nsmap))

        # ---------- 8. 其它所有容器（图形、公式等） ----------
        else:
            # 通通递归，保证不会遗漏任何可能的文本
            parts.append(deep_extract_text(child, nsmap))

        # 处理子元素后面的尾文本（通常为空白或分隔符）
        if child.tail:
            parts.append(child.tail)

    return ''.join(parts)


def get_header_text(header):
    """
    输入 python-docx 的 Header 对象，返回深度提取的全部文本。
    """
    if header is None:
        return ''
    return deep_extract_text(header._element)


def extract_all_headers_text(doc):
    """
    提取文档所有节、所有类型页眉的完整文本。
    返回字典：{ "section_X": { "default": "...", "first": "...", "even": "..." } }
    """
    result = {}
    for idx, section in enumerate(doc.sections):
        sec_label = f"section_{idx}"
        sec_headers = {}

        # 默认页眉（奇数页或通用）
        sec_headers['default'] = get_header_text(section.header)

        # 首页页眉
        if section.different_first_page_header_footer:
            sec_headers['first'] = get_header_text(section.first_page_header)
        else:
            sec_headers['first'] = ''

        # 偶数页页眉
        if section.different_odd_and_even_pages_header_footer:
            sec_headers['even'] = get_header_text(section.even_page_header)
        else:
            sec_headers['even'] = ''

        result[sec_label] = sec_headers
    return result


# =================== 使用示例 ===================
if __name__ == '__main__':
    doc = Document('复杂页眉.docx')   # 替换为你的文件

    all_headers = extract_all_headers_text(doc)
    for section_name, headers in all_headers.items():
        print(f"===== {section_name} =====")
        for header_type, text in headers.items():
            if text.strip():
                print(f"[{header_type}]\n{text}\n")
from docx import Document
from docx.oxml.ns import qn
from typing import Dict, Optional, Iterator, Tuple


class HeaderExtractor:
    """
    Word 文档页眉提取器
    
    提供面向对象的方式提取 Word 文档中的页眉内容，支持：
    - 提取所有节的页眉
    - 提取首页/奇数页/偶数页页眉
    - 支持多节文档
    """
    
    def __init__(self, doc_or_path):
        """
        初始化页眉提取器
        
        :param doc_or_path: Word 文档路径（字符串）或 Document 对象
        """
        if isinstance(doc_or_path, str):
            self._doc = Document(doc_or_path)
        elif hasattr(doc_or_path, 'sections') and hasattr(doc_or_path, 'paragraphs'):
            # 判断是否为 Document 对象（避免导入循环问题）
            self._doc = doc_or_path
        else:
            raise TypeError("参数必须是文档路径字符串或 Document 对象")
        
        self._headers_cache = None
    
    def _deep_extract_text(self, element, nsmap=None):
        """
        递归提取 OOXML 元素及其所有子元素中的文本
        """
        if nsmap is None:
            nsmap = {
                'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
                'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
                'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
            }
        
        parts = []
        
        # 注意：不处理 element.text，因为 lxml 的 text 属性会返回子元素中的文本
        # 文本只从 <w:t>、<w:instrText>、<w:delText> 等明确的文本元素中提取
        
        for child in element:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            
            # ---------- 直接文本节点 ----------
            if tag == 't':
                if child.text:
                    parts.append(child.text)
            elif tag == 'instrText':
                if child.text:
                    parts.append(child.text)
            elif tag == 'delText':
                if child.text:
                    parts.append(child.text)
            
            # ---------- 换行符 ----------
            elif tag == 'br':
                parts.append('\n')
            
            # ---------- 行内容器（run）----------
            elif tag == 'r':
                parts.append(self._deep_extract_text(child, nsmap))
            
            # ---------- 段落 ----------
            elif tag == 'p':
                parts.append(self._deep_extract_text(child, nsmap))
                parts.append('\n')
            
            # ---------- 表格 ----------
            elif tag == 'tbl':
                parts.append(self._deep_extract_text(child, nsmap))
            elif tag == 'tr':
                parts.append(self._deep_extract_text(child, nsmap))
                parts.append('\n')
            elif tag == 'tc':
                parts.append(self._deep_extract_text(child, nsmap))
                parts.append('\t')
            
            # ---------- 文本框 ----------
            elif tag == 'txbxContent':
                parts.append(self._deep_extract_text(child, nsmap))
            
            # ---------- 结构化文档标签 ----------
            elif tag == 'sdt':
                parts.append(self._deep_extract_text(child, nsmap))
            
            # ---------- 兼容性备用内容 ----------
            elif tag == 'AlternateContent':
                fallback = child.find(qn('mc:Fallback'))
                if fallback is not None:
                    parts.append(self._deep_extract_text(fallback, nsmap))
            
            # ---------- 其它所有容器 ----------
            else:
                parts.append(self._deep_extract_text(child, nsmap))
            
            # 处理子元素后面的尾文本
            if child.tail:
                parts.append(child.tail)
        
        return ''.join(parts)
    
    def _get_header_text(self, header) -> str:
        """
        提取单个页眉对象的文本内容
        """
        if header is None:
            return ''
        return self._deep_extract_text(header._element)
    
    def _has_different_odd_even_pages(self, section) -> bool:
        """
        检查节是否设置了奇偶页不同
        """
        sectPr = section._sectPr
        if sectPr is None:
            return False
        evenAndOddHeaders = sectPr.find(qn('w:evenAndOddHeaders'))
        return evenAndOddHeaders is not None
    
    def _extract_all_headers(self) -> Dict[str, Dict[str, str]]:
        """
        提取所有节的页眉内容（带缓存）
        """
        if self._headers_cache is not None:
            return self._headers_cache
        
        result = {}
        for idx, section in enumerate(self._doc.sections):
            sec_label = f"section_{idx}"
            sec_headers = {}
            
            # 默认页眉（奇数页或通用）
            sec_headers['default'] = self._get_header_text(section.header)
            
            # 首页页眉
            if section.different_first_page_header_footer:
                sec_headers['first'] = self._get_header_text(section.first_page_header)
            else:
                sec_headers['first'] = ''
            
            # 偶数页页眉
            if self._has_different_odd_even_pages(section):
                sec_headers['even'] = self._get_header_text(section.even_page_header)
            else:
                sec_headers['even'] = ''
            
            result[sec_label] = sec_headers
        
        self._headers_cache = result
        return result
    
    def get_all_headers(self) -> Dict[str, Dict[str, str]]:
        """
        获取所有节的页眉内容
        
        返回字典结构：
        {
            "section_0": {
                "default": "奇数页/通用页眉内容",
                "first": "首页页眉内容",
                "even": "偶数页页眉内容"
            },
            "section_1": { ... },
            ...
        }
        
        :return: 包含所有页眉内容的字典
        """
        return self._extract_all_headers()
    
    def get_section_headers(self, section_index: int = 0) -> Optional[Dict[str, str]]:
        """
        获取指定节的页眉内容
        
        :param section_index: 节索引，从 0 开始
        :return: 包含该节页眉内容的字典，节不存在时返回 None
        """
        all_headers = self._extract_all_headers()
        key = f"section_{section_index}"
        return all_headers.get(key)
    
    def get_header(self, section_index: int = 0, header_type: str = 'default') -> str:
        """
        获取指定节的指定类型页眉
        
        :param section_index: 节索引，从 0 开始
        :param header_type: 页眉类型，可选值：'default'（默认/奇数页）、'first'（首页）、'even'（偶数页）
        :return: 页眉文本内容，不存在时返回空字符串
        """
        section_headers = self.get_section_headers(section_index)
        if section_headers is None:
            return ''
        return section_headers.get(header_type, '')
    
    def get_default_header(self, section_index: int = 0) -> str:
        """
        获取指定节的默认（奇数页）页眉
        
        :param section_index: 节索引，从 0 开始
        :return: 页眉文本内容
        """
        return self.get_header(section_index, 'default')
    
    def get_first_page_header(self, section_index: int = 0) -> str:
        """
        获取指定节的首页页眉
        
        :param section_index: 节索引，从 0 开始
        :return: 页眉文本内容，未设置时返回空字符串
        """
        return self.get_header(section_index, 'first')
    
    def get_even_page_header(self, section_index: int = 0) -> str:
        """
        获取指定节的偶数页页眉
        
        :param section_index: 节索引，从 0 开始
        :return: 页眉文本内容，未设置时返回空字符串
        """
        return self.get_header(section_index, 'even')
    
    def section_count(self) -> int:
        """
        获取文档的节数量
        
        :return: 节数量
        """
        return len(self._doc.sections)
    
    def iter_sections(self) -> Iterator[Tuple[int, Dict[str, str]]]:
        """
        迭代所有节的页眉
        
        :return: 迭代器，返回 (节索引, 页眉字典)
        """
        all_headers = self._extract_all_headers()
        for idx in range(len(self._doc.sections)):
            key = f"section_{idx}"
            yield idx, all_headers.get(key, {})
    
    def __len__(self) -> int:
        """
        返回节数量
        """
        return self.section_count()
    
    def __getitem__(self, section_index: int) -> Dict[str, str]:
        """
        通过索引访问指定节的页眉
        
        :param section_index: 节索引
        :return: 页眉字典
        :raises IndexError: 索引越界时抛出
        """
        result = self.get_section_headers(section_index)
        if result is None:
            raise IndexError("节索引超出范围")
        return result
    
    def clear_cache(self):
        """
        清除缓存，下次提取时重新解析文档
        """
        self._headers_cache = None


# =================== 使用示例 ===================
if __name__ == '__main__':
    # 方式1：使用文档路径初始化
    extractor = HeaderExtractor('复杂页眉.docx')
    
    # 方式2：使用 Document 对象初始化
    # doc = Document('复杂页眉.docx')
    # extractor = HeaderExtractor(doc)
    
    print("===== 获取所有页眉 =====")
    all_headers = extractor.get_all_headers()
    for section_name, headers in all_headers.items():
        print(f"\n{section_name}:")
        for header_type, text in headers.items():
            if text.strip():
                print(f"  {header_type}: {repr(text.strip())}")
    
    print("\n===== 获取第一节的默认页眉 =====")
    print(extractor.get_default_header(0))
    
    print("\n===== 迭代所有节 =====")
    for idx, headers in extractor.iter_sections():
        print(f"\n第 {idx + 1} 节:")
        print(f"  默认页眉: {repr(headers.get('default', '').strip())}")
        print(f"  首页页眉: {repr(headers.get('first', '').strip())}")
        print(f"  偶数页页眉: {repr(headers.get('even', '').strip())}")
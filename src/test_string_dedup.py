import unittest
from string_dedup import deduplicate_longest, is_valid_pattern


class TestValidPattern(unittest.TestCase):
    """测试子串有效性检查函数"""
    
    def test_valid_pattern_with_spaces(self):
        """测试有效的子串（两边有空格）"""
        s = "Hi AB 25.0.0 测试"
        # "Hi" 是有效的（开头和空格之间）
        self.assertTrue(is_valid_pattern(s, 0, 2))
        # "AB" 是有效的（空格之间）
        self.assertTrue(is_valid_pattern(s, 3, 5))
        # "25.0.0" 是有效的（空格之间）
        self.assertTrue(is_valid_pattern(s, 6, 12))
        # "测试" 是有效的（空格和结尾之间）
        self.assertTrue(is_valid_pattern(s, 13, 15))
    
    def test_invalid_pattern_inside_word(self):
        """测试无效的子串（在单词内部）"""
        s = "Hi AB 25.0.0 测试"
        # ".0" 在 "25.0.0" 内部，不是有效的子串
        self.assertFalse(is_valid_pattern(s, 8, 10))
        # "5.0" 在 "25.0.0" 内部，不是有效的子串
        self.assertFalse(is_valid_pattern(s, 7, 10))
        # "B " 在 "AB " 内部，不是有效的子串
        self.assertFalse(is_valid_pattern(s, 4, 6))
    
    def test_invalid_pattern_all_spaces(self):
        """测试全空格子串（无效）"""
        s = "Hi   AB"
        # 全空格不是有效的子串
        self.assertFalse(is_valid_pattern(s, 2, 5))
    
    def test_single_word(self):
        """测试单个单词"""
        s = "Hello"
        # 整个字符串是有效的（开头和结尾）
        self.assertTrue(is_valid_pattern(s, 0, 5))
        # 部分字符串是无效的（在单词内部）
        self.assertFalse(is_valid_pattern(s, 1, 4))


class TestStringDedup(unittest.TestCase):
    """测试按照最大长度去重的字符串处理函数"""
    
    def test_no_duplicates_case(self):
        """测试无重复字符串（用户新需求）"""
        input_str = "Hi AB 25.0.0 测试"
        expected = "Hi AB 25.0.0 测试"
        self.assertEqual(deduplicate_longest(input_str), expected)
    
    def test_empty_string(self):
        """测试空字符串"""
        self.assertEqual(deduplicate_longest(""), "")
    
    def test_single_char(self):
        """测试单个字符"""
        self.assertEqual(deduplicate_longest("a"), "a")
    
    def test_single_word(self):
        """测试单个单词"""
        self.assertEqual(deduplicate_longest("Hello"), "Hello")
    
    def test_simple_duplicate(self):
        """测试简单重复（保留后面的空格）"""
        input_str = "Hi Hi"
        expected = "Hi "
        self.assertEqual(deduplicate_longest(input_str), expected)
    
    def test_longest_pattern_priority(self):
        """测试最长模式优先（保留分隔空格）"""
        input_str = "abc abc ab"
        expected = "abc  ab"
        self.assertEqual(deduplicate_longest(input_str), expected)
    
    def test_multiple_longest_patterns(self):
        """测试多个最长重复模式（保留分隔空格）"""
        input_str = "你好 你好 世界 世界"
        expected = "你好  世界 "
        self.assertEqual(deduplicate_longest(input_str), expected)
    
    def test_overlapping_patterns(self):
        """测试重叠模式（保留后面的空格）"""
        input_str = "aaa aaa"
        expected = "aaa "
        self.assertEqual(deduplicate_longest(input_str), expected)
    
    def test_whitespace_preservation(self):
        """测试空白字符保留"""
        input_str = "  hello   hello  "
        expected = "  hello  "
        self.assertEqual(deduplicate_longest(input_str), expected)
    
    def test_mixed_chars(self):
        """测试混合字符（循环去重，保留分隔空格）"""
        input_str = "abc123 abc123 xyzxyz"
        expected = "abc123  xyzxyz"
        self.assertEqual(deduplicate_longest(input_str), expected)
    
    def test_chinese_characters(self):
        """测试中文字符（保留分隔空格）"""
        input_str = "测试 测试 用例 用例"
        expected = "测试  用例 "
        self.assertEqual(deduplicate_longest(input_str), expected)
    
    def test_adjacent_duplicates(self):
        """测试相邻重复（保留分隔空格）"""
        input_str = "重复 重复 重复"
        expected = "重复  重复"
        self.assertEqual(deduplicate_longest(input_str), expected)
    
    def test_complex_case(self):
        """测试复杂场景（循环去重，保留分隔空格）"""
        input_str = "这是一个测试 这是一个测试 另一个测试 另一个测试"
        expected = "这是一个测试  另一个测试 "
        self.assertEqual(deduplicate_longest(input_str), expected)
    
    def test_pattern_with_spaces(self):
        """测试包含空格的模式（非相邻重复）"""
        input_str = "你真的 好 呀   你真的 好 呀"
        expected = "你真的 好 呀   "
        self.assertEqual(deduplicate_longest(input_str), expected)
    
    def test_new_case(self):
        """测试新用例（用户报告的问题，保留分隔空格）"""
        input_str = "Hi 25.0.0 测试.docx Hi 25.0.0 测试.docx 内部公开 内部公开"
        expected = "Hi 25.0.0 测试.docx  内部公开 "
        self.assertEqual(deduplicate_longest(input_str), expected)
    
    def test_no_duplicates_with_special_chars(self):
        """测试不含重复的特殊字符字符串"""
        input_str = "Hi AB 25.0.0 测试.docx"
        expected = "Hi AB 25.0.0 测试.docx"
        self.assertEqual(deduplicate_longest(input_str), expected)
    
    def test_duplicate_with_special_chars(self):
        """测试包含特殊字符的重复（保留分隔空格）"""
        input_str = "file.txt file.txt image.png image.png"
        expected = "file.txt  image.png "
        self.assertEqual(deduplicate_longest(input_str), expected)
    
    def test_partial_overlap_invalid(self):
        """测试部分重叠但无效的情况"""
        input_str = "abcdabc"
        expected = "abcdabc"
        self.assertEqual(deduplicate_longest(input_str), expected)
    
    def test_duplicate_at_boundary(self):
        """测试边界处的重复（保留后面的空格）"""
        input_str = "start end start end"
        expected = "start end "
        self.assertEqual(deduplicate_longest(input_str), expected)
    
    def test_leading_trailing_spaces(self):
        """测试前后有空格的情况（保留后面的空格）"""
        input_str = "  hello world  hello world  "
        expected = "  hello world  "
        self.assertEqual(deduplicate_longest(input_str), expected)
    
    def test_only_spaces(self):
        """测试只有空格的情况"""
        input_str = "   "
        expected = "   "
        self.assertEqual(deduplicate_longest(input_str), expected)
    
    def test_single_space_not_deduplicated(self):
        """测试单独空格不会被去重"""
        input_str = "a b c d"
        expected = "a b c d"
        self.assertEqual(deduplicate_longest(input_str), expected)
    
    def test_multiple_spaces_between_words(self):
        """测试单词间多个空格不会被去重"""
        input_str = "Hi   AB   25.0.0   测试"
        expected = "Hi   AB   25.0.0   测试"
        self.assertEqual(deduplicate_longest(input_str), expected)


if __name__ == "__main__":
    unittest.main(verbosity=2)
def is_valid_pattern(s: str, start: int, end: int) -> bool:
    """
    检查子串是否是有效的重复单元
    
    规则：
    1. 子串左边必须是空格或字符串开头
    2. 子串右边必须是空格或字符串结尾
    3. 子串不能全是空格
    
    参数:
        s: 原始字符串
        start: 子串起始位置
        end: 子串结束位置（不包含）
        
    返回:
        如果是有效子串返回True，否则返回False
    """
    n = len(s)
    
    # 规则3：子串不能全是空格
    pattern = s[start:end]
    if pattern.strip() == "":
        return False
    
    # 规则1：子串左边必须是空格或字符串开头
    if start > 0:
        left_char = s[start - 1]
        if left_char != ' ':
            return False
    
    # 规则2：子串右边必须是空格或字符串结尾
    if end < n:
        right_char = s[end]
        if right_char != ' ':
            return False
    
    return True


def deduplicate_longest(s: str) -> str:
    """
    按照最大长度将重复的字符串去重
    
    参数:
        s: 输入的字符串，可能包含重复的子字符串
        
    返回:
        去重后的字符串，保留最长的重复单元
        
    示例:
        输入: "你真的 好 呀   你真的 好 呀 内部公开 内部公开"
        输出: "你真的 好 呀  内部公开"
    """
    if not s:
        return s
    
    n = len(s)
    if n == 1:
        return s
    
    # 找到所有重复模式
    all_duplicates = []
    
    for length in range(n // 2, 0, -1):
        seen = {}
        for i in range(n - length + 1):
            end = i + length
            
            # 检查是否是有效子串
            if not is_valid_pattern(s, i, end):
                continue
            
            pattern = s[i:end]
            if pattern in seen:
                first_pos = seen[pattern]
                # 确保不重叠
                if i >= first_pos + length:
                    all_duplicates.append((-length, first_pos, i, length))
                    del seen[pattern]
            else:
                seen[pattern] = i
    
    if not all_duplicates:
        return s
    
    # 按长度降序、首次出现位置升序排序
    all_duplicates.sort()
    
    # 标记要删除的位置和已保留的位置
    to_remove = set()
    kept = set()
    
    for _, first_pos, second_pos, length in all_duplicates:
        # 检查是否与已保留或已删除的位置重叠
        conflict = False
        
        # 检查第一个模式是否与已删除位置重叠（不应删除第一个）
        for i in range(first_pos, first_pos + length):
            if i in to_remove:
                conflict = True
                break
        if conflict:
            continue
        
        # 检查第二个模式是否与已保留位置重叠
        for i in range(second_pos, second_pos + length):
            if i in kept:
                conflict = True
                break
        if conflict:
            continue
        
        # 检查第二个模式是否与已删除位置重叠
        has_deleted = False
        for i in range(second_pos, second_pos + length):
            if i in to_remove:
                has_deleted = True
                break
        
        # 标记第一个模式为保留
        for i in range(first_pos, first_pos + length):
            kept.add(i)
        
        # 标记第二个模式为删除（如果还没被删除）
        if not has_deleted:
            for i in range(second_pos, second_pos + length):
                to_remove.add(i)
    
    # 构建结果
    result = []
    for i in range(n):
        if i not in to_remove:
            result.append(s[i])
    
    return ''.join(result)


def main():
    """主函数，演示字符串去重功能"""
    # 示例输入
    input_str = "Hi AB 25.0.0 测试"
    input_str = "Hi AB 25.0.0 测试  Hi AB 25.0.0 ABAB 测试"


    print("=== 字符串去重功能演示 ===")
    print(f"输入字符串: {repr(input_str)}")
    
    # 调用去重函数
    result = deduplicate_longest(input_str)
    
    print(f"去重结果: {repr(result)}")
    print("=========================")


if __name__ == "__main__":
    main()
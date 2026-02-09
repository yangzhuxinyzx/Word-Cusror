import re
import os
import subprocess
import sys
from typing import Dict, List
from pathlib import Path
import cn2an

def process_markdown_file(input_file: str = 'index.md', output_file: str = 'output.md') -> None:
    # Create output directory if it doesn't exist
    output_dir = Path('./output')
    output_dir.mkdir(exist_ok=True)
    
    if not os.path.exists(input_file):
        print(f"错误：输入文件 {input_file} 不存在")
        return
    
    with open(input_file, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    chapter_num = 0  # 二级标题章节号
    section_nums = [0] * 4  # 三级、四级、五级、六级标题的编号
    img_count = 0  # 当前章节图片计数
    table_count = 0  # 当前章节表格计数
    skipped_chapters = 0  # 已跳过的二级标题计数，用于跳过摘要
    processed_lines: List[str] = []
    in_code_block = False  # 标记是否在代码块中
    
    for line in lines:
        # 检测代码块开始或结束
        if line.strip().startswith('```'):
            in_code_block = not in_code_block
            processed_lines.append(line)
            continue
            
        # 如果在代码块中，直接添加行不处理
        if in_code_block:
            processed_lines.append(line)
            continue

        # 处理二级标题 (##)
        if line.startswith('## '):
            skipped_chapters += 1
            if skipped_chapters <= 2:
                # 跳过前两个二级标题，不做编号处理，因为这俩是摘要
                # 如果要改实现，记得改这里
                processed_lines.append(line)
                continue
                
            chapter_num += 1
            section_nums = [0] * 4
            img_count = 0
            table_count = 0
            processed_line = f'## 第{cn2an.an2cn(chapter_num)}章 {line[3:]}'
            processed_lines.append(processed_line)
            continue
        
        elif line.startswith('### '):
            if skipped_chapters <= 2:
                processed_lines.append(line)
                continue
                
            section_nums[0] += 1
            section_nums[1:] = [0] * 3
            processed_line = f'### {chapter_num}.{section_nums[0]} {line[4:]}'
            processed_lines.append(processed_line)
            continue
        
        elif line.startswith('#### '):
            if skipped_chapters <= 2:
                processed_lines.append(line)
                continue
                
            section_nums[1] += 1
            section_nums[2:] = [0] * 2
            processed_line = f'#### {chapter_num}.{section_nums[0]}.{section_nums[1]} {line[5:]}'
            processed_lines.append(processed_line)
            continue
        
        elif line.startswith('##### '):
            if skipped_chapters <= 2:
                processed_lines.append(line)
                continue
                
            section_nums[2] += 1
            section_nums[3] = 0
            processed_line = f'##### {chapter_num}.{section_nums[0]}.{section_nums[1]}.{section_nums[2]} {line[6:]}'
            processed_lines.append(processed_line)
            continue
        
        elif line.startswith('###### '):
            if skipped_chapters <= 2:
                processed_lines.append(line)
                continue
                
            section_nums[3] += 1
            processed_line = f'###### {chapter_num}.{section_nums[0]}.{section_nums[1]}.{section_nums[2]}.{section_nums[3]} {line[7:]}'
            processed_lines.append(processed_line)
            continue
        
        img_inline_match = re.match(r'^!\[(.*?)\]\((.*?)\)', line)  # 处理内联图片格式 ![alt](path)
        img_reference_match = re.match(r'^!\[(.*?)\]\[(.*?)\]', line)  # 处理引用格式 ![alt][id]
        reference_def_match = re.match(r'^\[(.*?)\]:\s*(.*?)(?:\s+\{(.*?)\})?$', line)  # 处理引用定义 [id]: path {attrs}

        if img_inline_match:
            if skipped_chapters <= 2:
                processed_lines.append(line)
                continue
                
            img_count += 1
            caption = img_inline_match.group(1)
            path = img_inline_match.group(2)
            new_caption = f'图{chapter_num}-{img_count} {caption}'
            processed_line = f'![{new_caption}]({path})\n'
            processed_lines.append(processed_line)
            continue
        elif img_reference_match:
            if skipped_chapters <= 2:
                processed_lines.append(line)
                continue
                
            img_count += 1
            caption = img_reference_match.group(1)
            ref_id = img_reference_match.group(2)
            new_caption = f'图{chapter_num}-{img_count} {caption}'
            processed_line = f'![{new_caption}][{ref_id}]\n'
            processed_lines.append(processed_line)
            continue
        elif reference_def_match:
            # 保留引用定义不变
            processed_lines.append(line)
            continue
        
        if line.startswith('Table: '):
            if skipped_chapters <= 2:
                processed_lines.append(line)
                continue
                
            table_count += 1
            caption = line[7:].strip()
            processed_line = f'Table: 表{chapter_num}-{table_count} {caption}\n'
            processed_lines.append(processed_line)
            continue
        
        processed_lines.append(line)
    
    # Write the processed markdown file
    with open(output_file, 'w', encoding='utf-8') as f:
        f.writelines(processed_lines)
    print(f"处理完成，Markdown 结果已保存到 {output_file}")
    
    # Delete existing output.docx if it exists
    docx_output = output_dir / 'output.docx'
    if docx_output.exists():
        try:
            docx_output.unlink()
            print(f"已删除已存在的 {docx_output}")
        except Exception as e:
            print(f"删除 {docx_output} 时出错：{e}")
            return
    
    # Run pandoc command
    pandoc_cmd = [
        'pandoc',
        output_file,
        '--lua-filter=./pandocScripts/filter.lua',
        '--bibliography=references.bib',
        '--citeproc',
        '--csl=./pandocAssets/chinese-gb7714-2005-numeric.csl',
        '--reference-doc=./pandocAssets/custom-reference.docx',
        '-o',
        str(docx_output)
    ]
    
    print("正在使用 pandoc 生成 Word 文档...")
    try:
        subprocess.run(pandoc_cmd, check=True)
        print(f"Word 文档已成功生成到 {docx_output}")
    except subprocess.CalledProcessError as e:
        print(f"pandoc 命令执行失败：{e}")
        return
    except FileNotFoundError:
        print("未找到 pandoc，请确保 pandoc 已安装并添加到 PATH 环境变量中")
        return
    
    # Delete the intermediate markdown file
    try:
        os.remove(output_file)
        print(f"已删除临时文件 {output_file}")
    except Exception as e:
        print(f"删除 {output_file} 时出错：{e}")
    
    # Open the output.docx with system default application
    if sys.platform == 'win32':
        os.startfile(docx_output)

if __name__ == '__main__':
    process_markdown_file()

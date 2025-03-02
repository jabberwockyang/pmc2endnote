#!/usr/bin/env python3
"""
PMC文章处理器 - 一站式解决方案

这个脚本接受PMID作为输入，下载PMC文章，并生成Word文档和RIS文件。
使用方法:
    python pmc_pro.py --pmid 32066954
    python pmc_pro.py --pmids 32066954,31978945,32376714
    python pmc_pro.py --file pmids.txt
"""

import os
import sys
import argparse
import json
import shutil
from pathlib import Path
import importlib.util
from loguru import logger

def import_module_from_file(module_name, file_path):
    """从文件路径导入模块"""
    spec = importlib.util.spec_from_file_location(module_name, file_path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module

def setup_logger():
    """配置日志记录器"""
    logger.remove()  # 移除默认处理程序
    logger.add(sys.stderr, format="{time:YYYY-MM-DD HH:mm:ss} | {level} | {message}", level="INFO")
    logger.add("pmc_pro.log", rotation="10 MB", level="DEBUG")

def parse_arguments():
    """解析命令行参数"""
    parser = argparse.ArgumentParser(description="PMC文章处理器 - 下载PMC文章并生成Word和RIS文件")
    
    # 创建互斥组，确保只使用一种输入方式
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--pmid", help="单个PMID")
    group.add_argument("--pmids", help="多个PMID，用逗号分隔")
    group.add_argument("--file", help="包含PMID列表的文件，每行一个PMID")
    
    parser.add_argument("--output", "-o", default="pmc_output", help="输出目录")
    parser.add_argument("--keep-xml", action="store_true", help="保留中间XML文件")
    parser.add_argument("--verbose", "-v", action="store_true", help="显示详细输出")
    
    return parser.parse_args()

def get_pmids(args):
    """从命令行参数获取PMID列表"""
    if args.pmid:
        return [args.pmid]
    elif args.pmids:
        return [pmid.strip() for pmid in args.pmids.split(",")]
    elif args.file:
        with open(args.file, "r") as f:
            return [line.strip() for line in f if line.strip()]
    return []

def process_pmid(pmid, output_dir, pmc_converter, generate_ris, generate_word):
    """处理单个PMID"""
    logger.info(f"处理PMID: {pmid}")
    
    # 创建临时目录
    temp_dir = os.path.join(output_dir, "temp")
    os.makedirs(temp_dir, exist_ok=True)
    
    try:
        # 步骤1: 使用PMC_converter下载文章
        logger.info("下载PMC文章...")
        article_finder = pmc_converter.ArticleRetrieval(pmids=[pmid], repo_dir=temp_dir)
        article_finder.initiallize()
        
        # 查找下载的XML文件
        xml_files = [f for f in os.listdir(temp_dir) if f.endswith(".xml")]
        if not xml_files:
            logger.error(f"未找到PMID {pmid}的XML文件")
            return None
        
        xml_file = os.path.join(temp_dir, xml_files[0])
        logger.info(f"找到XML文件: {xml_file}")
        
        # 提取PMCID
        tree = pmc_converter.ET.parse(xml_file)
        root = tree.getroot()
        pmcid = None
        article_id = root.find('.//article-id[@pub-id-type="pmc"]')
        if article_id is not None and article_id.text:
            pmcid = f"PMC{article_id.text}"
        else:
            logger.warning(f"无法从XML提取PMCID，使用PMID作为标识符")
            pmcid = f"PMID{pmid}"
        
        logger.info(f"提取的PMCID: {pmcid}")
        
        # 步骤2: 提取参考文献并生成RIS文件
        logger.info("提取参考文献...")
        ris_file, references, extracted_pmcid = generate_ris.extract_references_to_ris(xml_file)
        
        # 使用提取的PMCID（如果可用）
        if extracted_pmcid:
            pmcid = extracted_pmcid
            logger.info(f"使用从RIS提取的PMCID: {pmcid}")
        
        # 步骤3: 创建Word文档
        logger.info("创建Word文档...")
        docx_file = generate_word.create_word_with_citation_markers(xml_file, references, pmcid)
        
        # 步骤4: 将文件移动到以PMCID命名的目录
        pmcid_dir = os.path.join(output_dir, pmcid)
        os.makedirs(pmcid_dir, exist_ok=True)
        
        # 复制文件到最终目录
        final_xml = os.path.join(pmcid_dir, f"{pmcid}.xml")
        final_ris = os.path.join(pmcid_dir, f"{pmcid}.ris")
        final_docx = os.path.join(pmcid_dir, f"{pmcid}.docx")
        
        shutil.copy2(xml_file, final_xml)
        shutil.copy2(ris_file, final_ris)
        shutil.copy2(docx_file, final_docx)
        
        logger.info(f"文件已保存到目录: {pmcid_dir}")
        logger.info(f"XML: {final_xml}")
        logger.info(f"RIS: {final_ris}")
        logger.info(f"Word: {final_docx}")
        
        # 步骤5: 追加到主RIS文件
        import_dir, master_ris_path = setup_endnote_import_folder(output_dir)
        if append_to_master_ris(final_ris, master_ris_path):
            logger.info(f"已将引用追加到主RIS文件: {master_ris_path}")
        
        return {
            "pmid": pmid,
            "pmcid": pmcid,
            "xml": final_xml,
            "ris": final_ris,
            "docx": final_docx
        }
    
    except Exception as e:
        logger.error(f"处理PMID {pmid}时出错: {str(e)}")
        import traceback
        logger.debug(traceback.format_exc())
        return None
    finally:
        # 清理临时目录
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)

def append_to_master_ris(ris_file, master_ris_path):
    """将RIS文件内容追加到主RIS文件中，并清理重复记录"""
    # 读取新RIS文件内容
    with open(ris_file, 'r', encoding='utf-8') as f:
        new_content = f.read().strip()
    
    # 确保内容以ER  - 结束
    if not new_content.endswith('ER  - '):
        new_content += '\nER  - '
    
    # 检查主RIS文件是否存在
    if not os.path.exists(master_ris_path):
        # 如果不存在，直接写入新内容
        with open(master_ris_path, 'w', encoding='utf-8') as f:
            f.write(new_content)
        logger.info(f"创建主RIS文件: {master_ris_path}")
        return True
    
    # 读取主RIS文件内容
    with open(master_ris_path, 'r', encoding='utf-8') as f:
        master_content = f.read().strip()
    
    # 确保主内容以ER  - 结束
    if not master_content.endswith('ER  - '):
        master_content += '\nER  - '
    
    # 追加新内容，确保有空行分隔
    with open(master_ris_path, 'w', encoding='utf-8') as f:
        f.write(master_content + '\n\n' + new_content)
    
    logger.info(f"已将新内容追加到主RIS文件: {master_ris_path}")
    
    # 清理重复记录
    clean_master_ris(master_ris_path)
    
    return True

def setup_endnote_import_folder(output_dir):
    """设置EndNote导入文件夹和主RIS文件"""
    import_dir = os.path.join(output_dir, "endnote_import")
    os.makedirs(import_dir, exist_ok=True)
    
    # 创建主RIS文件路径
    master_ris_path = os.path.join(import_dir, "master_references.ris")
    
    # 创建说明文件
    readme_path = os.path.join(import_dir, "README.txt")
    with open(readme_path, "w", encoding='utf-8') as f:
        f.write("EndNote导入文件夹\n\n")
        f.write("主RIS文件: master_references.ris\n")
        f.write("此文件包含所有处理过的文献的引用信息。\n\n")
        f.write("使用方法:\n")
        f.write("1. 打开EndNote\n")
        f.write("2. 选择 File > Import > File...\n")
        f.write("3. 选择 master_references.ris 文件\n")
        f.write("4. Import Option 选择 'Reference Manager (RIS)'\n")
        f.write("5. 点击 Import 按钮\n\n")
        f.write("每次处理新的文献后，master_references.ris 文件会自动更新。\n")
        f.write("您只需重新导入此文件即可获取所有新增的引用。\n")
    
    return import_dir, master_ris_path

def clean_master_ris(master_ris_path):
    """
    检查主RIS文件中的重复记录，保留每个唯一LB标识符的最新记录
    
    参数:
        master_ris_path: 主RIS文件的路径
    
    返回:
        bool: 是否成功清理文件
    """
    if not os.path.exists(master_ris_path):
        logger.warning(f"主RIS文件不存在: {master_ris_path}")
        return False
    
    try:
        # 读取主RIS文件内容
        with open(master_ris_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 按记录分割内容（每条记录以ER  - 结束）
        records_raw = content.split('ER  - ')
        records = []
        
        # 处理每条记录，提取LB标识符
        for record in records_raw:
            record = record.strip()
            if not record:  # 跳过空记录
                continue
                
            # 确保记录以ER  - 结束
            record_complete = record + '\nER  - '
            
            # 提取LB标识符
            lb_identifier = None
            for line in record.split('\n'):
                if line.startswith('LB  - ^'):
                    lb_identifier = line.strip()
                    break
            
            if lb_identifier:
                records.append((lb_identifier, record_complete))
            else:
                # 没有LB标识符的记录也保留
                records.append((f"NO_LB_{len(records)}", record_complete))
        
        # 检查是否有重复的LB标识符
        lb_dict = {}
        duplicates_found = False
        
        for lb, record in records:
            if lb in lb_dict:
                duplicates_found = True
                logger.warning(f"发现重复的LB标识符: {lb}")
            # 总是保存最新的记录（列表中靠后的记录）
            lb_dict[lb] = record
        
        # 如果发现重复，重写文件
        if duplicates_found:
            logger.info(f"发现重复记录，正在清理主RIS文件...")
            
            # 按原始顺序重建内容（保持最新的记录）
            unique_records = []
            seen_lb = set()
            
            # 反向遍历以保留最新的记录
            for lb, record in reversed(records):
                if lb not in seen_lb:
                    unique_records.append(record)
                    seen_lb.add(lb)
            
            # 恢复原始顺序
            unique_records.reverse()
            
            # 写入清理后的内容
            with open(master_ris_path, 'w', encoding='utf-8') as f:
                f.write('\n\n'.join(unique_records))
            
            logger.info(f"主RIS文件已清理，移除了 {len(records) - len(unique_records)} 条重复记录")
            return True
        else:
            logger.info(f"主RIS文件中没有发现重复记录")
            return False
    
    except Exception as e:
        logger.error(f"清理主RIS文件时出错: {str(e)}")
        import traceback
        logger.debug(traceback.format_exc())
        return False

def main():
    """主函数"""
    # 设置日志
    setup_logger()
    
    # 解析命令行参数
    args = parse_arguments()
    
    # 获取PMID列表
    pmids = get_pmids(args)
    logger.info(f"处理 {len(pmids)} 个PMID: {', '.join(pmids)}")
    
    # 创建输出目录
    output_dir = args.output
    os.makedirs(output_dir, exist_ok=True)
    
    # 导入所需模块
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 导入PMC_converter.py
    pmc_converter_path = os.path.join(script_dir, "PMC_converter.py")
    if not os.path.exists(pmc_converter_path):
        logger.error(f"找不到PMC_converter.py: {pmc_converter_path}")
        sys.exit(1)
    pmc_converter = import_module_from_file("pmc_converter", pmc_converter_path)
    
    # 导入generate_ris.py
    generate_ris_path = os.path.join(script_dir, "generate_ris.py")
    if not os.path.exists(generate_ris_path):
        logger.error(f"找不到generate_ris.py: {generate_ris_path}")
        sys.exit(1)
    generate_ris = import_module_from_file("generate_ris", generate_ris_path)
    
    # 导入generate_word.py
    generate_word_path = os.path.join(script_dir, "generate_word.py")
    if not os.path.exists(generate_word_path):
        logger.error(f"找不到generate_word.py: {generate_word_path}")
        sys.exit(1)
    generate_word = import_module_from_file("generate_word", generate_word_path)
    
    # 处理每个PMID
    results = []
    for pmid in pmids:
        result = process_pmid(pmid, output_dir, pmc_converter, generate_ris, generate_word)
        if result:
            results.append(result)
    
    # 保存处理结果
    results_file = os.path.join(output_dir, "processing_results.json")
    with open(results_file, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    
    logger.info(f"处理完成! 成功处理 {len(results)}/{len(pmids)} 个PMID")
    logger.info(f"结果摘要已保存到: {results_file}")
    
    # 显示使用说明
    print("\n使用说明:")
    print("1. 所有引用已合并到主RIS文件:")
    print(f"   {os.path.join(output_dir, 'endnote_import', 'master_references.ris')}")
    print("2. 在EndNote中导入此RIS文件:")
    print("   File > Import > File... > 选择主RIS文件 > Import Option选择'Reference Manager (RIS)'")
    print("3. 确认Label字段包含带有PMCID的唯一标识符")
    print("4. 在EndNote偏好设置中启用Label匹配")
    print("5. 打开生成的Word文档")
    print("6. 使用EndNote的'Update Citations and Bibliography'功能")
    print("\n每次处理新的文献后，只需重新导入主RIS文件即可更新EndNote库。")

if __name__ == "__main__":
    main()
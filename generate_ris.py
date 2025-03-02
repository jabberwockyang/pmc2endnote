import xml.etree.ElementTree as ET
import os

def get_element_text(element):
    """
    从XML元素中提取所有文本内容，包括子元素中的文本
    
    参数:
        element: XML元素对象
    
    返回:
        元素及其子元素中的所有文本内容，作为一个字符串
    """
    if element is None:
        return ""
    
    # 获取元素的直接文本内容
    text = element.text or ""
    
    # 递归获取所有子元素的文本内容
    for child in element:
        # 添加子元素的文本
        text += get_element_text(child)
        # 添加子元素后的尾随文本
        if child.tail:
            text += child.tail
    
    # 清理文本（移除多余空格、换行符等）
    return " ".join(text.split())

def extract_references_to_ris(xml_path):
    """从PMC XML文件提取参考文献并生成RIS文件，使用PMCID创建唯一标识符"""
    tree = ET.parse(xml_path)
    root = tree.getroot()
    references = []
    
    # 提取PMCID - 方法1：从文件名提取
    file_name = os.path.basename(xml_path)
    pmcid = ""
    if file_name.startswith("PMC") and file_name.endswith(".xml"):
        pmcid = file_name[:-4]  # 移除.xml后缀
    
    # 提取PMCID - 方法2：从XML内容提取
    if not pmcid:
        article_ids = root.findall(".//article-id")
        for article_id in article_ids:
            if article_id.get("pub-id-type") == "pmc":
                pmcid = "PMC" + article_id.text
                break
    
    # 如果仍未找到PMCID，使用默认值
    if not pmcid:
        pmcid = "PMC_UNKNOWN"
        print("警告: 无法从文件名或内容中提取PMCID，使用默认值")
    
    # 查找所有参考文献
    ref_elements = root.findall(".//ref")
    
    for i, ref in enumerate(ref_elements):
        ref_id = ref.get("id", f"ref{i+1}")
        
        # 提取引用类型
        pub_type = ""
        mixed_citation = ref.find(".//mixed-citation")
        if mixed_citation is not None:
            pub_type = mixed_citation.get("publication-type", "")
        
        # 确定RIS类型代码
        ris_type = "JOUR"  # 默认为期刊文章
        if pub_type == "book":
            ris_type = "BOOK"
        elif pub_type == "chapter":
            ris_type = "CHAP"
        elif pub_type == "conference":
            ris_type = "CONF"
        elif pub_type == "web":
            ris_type = "ELEC"
        
        # 提取标题
        title = ""
        article_title = ref.find(".//article-title")
        if article_title is not None:
            title = get_element_text(article_title)
        
        # 提取作者信息 - 这是新添加的部分
        authors = []
        person_groups = ref.findall(".//person-group")
        for person_group in person_groups:
            # 检查是否是作者组
            if person_group.get("person-group-type") in ["author", None]:
                # 提取所有作者
                for name_elem in person_group.findall(".//string-name"):
                    surname = name_elem.find("surname")
                    given_names = name_elem.find("given-names")
                    
                    if surname is not None:
                        surname_text = get_element_text(surname)
                        given_names_text = get_element_text(given_names) if given_names is not None else ""
                        
                        # 格式化作者名: 姓, 名
                        if given_names_text:
                            author = f"{surname_text}, {given_names_text}"
                        else:
                            author = surname_text
                        
                        authors.append(author)
        
        # 提取期刊名称
        journal = ""
        source = ref.find(".//source")
        if source is not None:
            journal = get_element_text(source)
        
        # 提取年份
        year = ""
        year_elem = ref.find(".//year")
        if year_elem is not None:
            year = get_element_text(year_elem)
        
        # 提取卷号
        volume = ""
        volume_elem = ref.find(".//volume")
        if volume_elem is not None:
            volume = get_element_text(volume_elem)
        
        # 提取期号
        issue = ""
        issue_elem = ref.find(".//issue")
        if issue_elem is not None:
            issue = get_element_text(issue_elem)
        
        # 提取页码
        pages = ""
        fpage = ref.find(".//fpage")
        lpage = ref.find(".//lpage")
        if fpage is not None:
            pages = get_element_text(fpage)
            if lpage is not None:
                pages += "-" + get_element_text(lpage)
        
        # 提取DOI
        doi = ""
        pub_id = ref.find(".//pub-id[@pub-id-type='doi']")
        if pub_id is not None:
            doi = get_element_text(pub_id)
        
        # 提取PMID
        pmid = ""
        pub_id = ref.find(".//pub-id[@pub-id-type='pmid']")
        if pub_id is not None:
            pmid = get_element_text(pub_id)
        
        # 创建唯一ID
        unique_id = f"{pmcid}_{ref_id}"
        
        # 将提取的信息添加到引用列表
        references.append({
            "id": ref_id,
            "unique_id": unique_id,
            "type": ris_type,
            "title": title,
            "authors": authors,  # 新添加的作者列表
            "journal": journal,
            "year": year,
            "volume": volume,
            "issue": issue,
            "pages": pages,
            "doi": doi,
            "pmid": pmid
        })
    
    # 生成RIS文件
    ris_path = os.path.splitext(xml_path)[0] + ".ris"
    with open(ris_path, "w", encoding="utf-8") as f:
        for ref in references:
            # 写入记录类型
            f.write(f"TY  - {ref['type']}\n")
            
            # 写入标题
            if ref['title']:
                f.write(f"TI  - {ref['title']}\n")
            
            # 写入作者 - 这是新添加的部分
            for author in ref['authors']:
                f.write(f"AU  - {author}\n")
            
            # 写入期刊名称
            if ref['journal']:
                f.write(f"JO  - {ref['journal']}\n")
            
            # 写入年份
            if ref['year']:
                f.write(f"PY  - {ref['year']}\n")
            
            # 写入卷号
            if ref['volume']:
                f.write(f"VL  - {ref['volume']}\n")
            
            # 写入期号
            if ref['issue']:
                f.write(f"IS  - {ref['issue']}\n")
            
            # 写入页码
            if ref['pages']:
                f.write(f"SP  - {ref['pages']}\n")
            
            # 写入DOI
            if ref['doi']:
                f.write(f"DO  - {ref['doi']}\n")
            
            # 写入PMID
            if ref['pmid']:
                f.write(f"AN  - {ref['pmid']}\n")
            
            # 写入带PMCID的唯一Label
            f.write(f"LB  - ^{ref['unique_id']}$\n")
            
            # 结束当前记录
            f.write("ER  - \n\n")
    
    print(f"已提取 {len(references)} 条参考文献并保存到 {ris_path}")
    return ris_path, references, pmcid

if __name__ == "__main__":
    xml_file = "repodir/PMC7096285.xml"  # 替换为您的XML文件路径
    ris_file, refs, pmcid = extract_references_to_ris(xml_file)
    print(f"RIS文件已生成: {ris_file}")
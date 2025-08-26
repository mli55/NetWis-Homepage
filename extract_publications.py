#!/usr/bin/env python3
"""
简单的论文提取工具 - 将publications.html中的论文提取到Excel文件
"""

import pandas as pd
import re
from bs4 import BeautifulSoup
import os
from datetime import datetime

def extract_publications_to_excel():
    """从publications.html提取论文到Excel文件"""
    
    html_file = "publications.html"
    excel_file = f"NetWIS_Publications_{datetime.now().strftime('%Y%m%d')}.xlsx"
    
    print(f"📚 正在从 {html_file} 提取论文...")
    
    try:
        # 读取HTML文件
        with open(html_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        soup = BeautifulSoup(content, 'html.parser')
        publications = []
        
        # 查找所有年份标题
        year_headers = soup.find_all('h2', class_='year-header')
        
        for year_header in year_headers:
            year = year_header.text.strip()
            print(f"📅 处理 {year} 年的论文...")
            
            # 查找这个年份后面的所有论文
            current_element = year_header.next_sibling
            
            while current_element:
                if current_element.name == 'div' and current_element.get('class') and 'publication-row' in current_element.get('class'):
                    pub_details = current_element.find('div', class_='publication-details')
                    
                    if pub_details:
                        # 提取论文信息
                        title_elem = pub_details.find('div', class_='pub-title')
                        authors_elem = pub_details.find('div', class_='pub-authors')
                        venue_elem = pub_details.find('div', class_='pub-venue')
                        id_elem = pub_details.find('div', class_='pub-id')
                        
                        # 获取各字段内容
                        title = title_elem.text.strip() if title_elem else "无标题"
                        authors = authors_elem.text.strip() if authors_elem else "无作者"
                        venue = venue_elem.text.strip() if venue_elem else "无venue"
                        
                        # 提取论文ID和PDF链接
                        pub_id = ""
                        pdf_link = ""
                        pub_type = ""
                        
                        if id_elem:
                            id_text = id_elem.text.strip()
                            # 提取ID模式 [JP-50] 或 [CP-125]
                            id_match = re.search(r'\[(.*?)\]', id_text)
                            if id_match:
                                pub_id = id_match.group(1)
                                # 判断论文类型
                                if pub_id.startswith('JP'):
                                    pub_type = "期刊论文"
                                elif pub_id.startswith('CP'):
                                    pub_type = "会议论文"
                                elif pub_id.startswith('WP'):
                                    pub_type = "研讨会论文"
                                else:
                                    pub_type = "其他"
                            
                            # 提取PDF链接
                            pdf_a = id_elem.find('a', class_='pdf-link')
                            if pdf_a and pdf_a.get('href'):
                                pdf_link = pdf_a.get('href')
                        
                        publications.append({
                            '年份': year,
                            '标题': title,
                            '作者': authors,
                            '发表venue': venue,
                            '论文ID': pub_id,
                            '论文类型': pub_type,
                            'PDF链接': pdf_link if pdf_link else "无",
                        })
                
                elif current_element.name == 'h2' and current_element.get('class') and 'year-header' in current_element.get('class'):
                    # 遇到下一个年份标题，跳出
                    break
                
                current_element = current_element.next_sibling
        
        # 创建DataFrame并保存到Excel
        if publications:
            df = pd.DataFrame(publications)
            # 按年份降序，然后按论文ID排序
            df = df.sort_values(['年份', '论文ID'], ascending=[False, True])
            
            # 保存到Excel
            with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='论文列表', index=False)
                
                # 调整列宽
                worksheet = writer.sheets['论文列表']
                column_widths = {
                    'A': 8,   # 年份
                    'B': 60,  # 标题
                    'C': 30,  # 作者
                    'D': 70,  # venue
                    'E': 12,  # 论文ID
                    'F': 12,  # 论文类型
                    'G': 15   # PDF链接
                }
                
                for column, width in column_widths.items():
                    worksheet.column_dimensions[column].width = width
            
            print(f"\n✅ 成功提取 {len(publications)} 篇论文到 {excel_file}")
            
            # 打印统计信息
            print(f"\n📊 统计信息:")
            print(f"   总论文数: {len(publications)}")
            print(f"   年份范围: {df['年份'].min()} - {df['年份'].max()}")
            
            # 按年份统计
            year_counts = df['年份'].value_counts().sort_index(ascending=False)
            print(f"\n📅 各年份论文数量:")
            for year, count in year_counts.items():
                print(f"   {year}: {count}篇")
            
            # 按类型统计
            if '论文类型' in df.columns:
                type_counts = df['论文类型'].value_counts()
                print(f"\n📝 论文类型统计:")
                for ptype, count in type_counts.items():
                    if ptype:  # 忽略空值
                        print(f"   {ptype}: {count}篇")
            
            return excel_file
        
        else:
            print("❌ 未找到任何论文数据")
            return None
    
    except Exception as e:
        print(f"❌ 提取过程中出错: {str(e)}")
        return None

def main():
    print("🔬 NetWIS Lab 论文提取工具")
    print("=" * 40)
    
    if not os.path.exists("publications.html"):
        print("❌ 找不到 publications.html 文件")
        print("💡 请确保在正确的目录中运行此脚本")
        return
    
    excel_file = extract_publications_to_excel()
    
    if excel_file:
        print(f"\n🎉 提取完成！")
        print(f"📁 Excel文件已保存为: {excel_file}")
        print(f"💡 您现在可以用Excel打开此文件查看和编辑论文数据")
    else:
        print(f"\n❌ 提取失败，请检查HTML文件格式")

if __name__ == "__main__":
    main()

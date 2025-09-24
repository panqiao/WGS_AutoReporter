#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import numpy as np

def preprocess_phenotypes(input_file, output_file):
    """
    预处理Excel文件，将包含多个phenotypes的行拆分成独立的行
    
    参数:
    input_file: 原始Excel文件路径
    output_file: 处理后的Excel文件路径
    """
    
    print("读取原始Excel文件...")
    df = pd.read_excel(input_file)
    print(f"原始数据: {len(df)} 行")
    
    # 创建新的数据列表
    new_rows = []
    
    # 统计信息
    split_count = 0
    multi_phenotype_rows = 0
    
    # 遍历每一行
    for idx, row in df.iterrows():
        phenotype_value = row['Phenotype(s)']
        
        # 检查是否为空值
        if pd.isna(phenotype_value):
            # 保留空值行
            new_rows.append(row.to_dict())
            continue
        
        # 转换为字符串
        phenotype_str = str(phenotype_value)
        
        # 检查是否包含分号
        if ';' in phenotype_str:
            multi_phenotype_rows += 1
            # 拆分phenotypes
            phenotypes = [p.strip() for p in phenotype_str.split(';')]
            split_count += len(phenotypes) - 1  # 减1是因为原始行算1个
            
            # 为每个phenotype创建新行
            for phenotype in phenotypes:
                new_row = row.to_dict()  # 复制整行数据
                new_row['Phenotype(s)'] = phenotype  # 更新Phenotype列
                new_rows.append(new_row)
        else:
            # 不包含分号的行直接添加
            new_rows.append(row.to_dict())
    
    # 创建新的DataFrame
    new_df = pd.DataFrame(new_rows)
    
    print(f"\n处理统计:")
    print(f"  - 包含多个phenotypes的行: {multi_phenotype_rows}")
    print(f"  - 拆分产生的新行: {split_count}")
    print(f"  - 处理后总行数: {len(new_df)} 行")
    
    # 保存到新的Excel文件
    new_df.to_excel(output_file, index=False)
    print(f"\n预处理后的文件已保存至: {output_file}")
    
    # 显示拆分示例
    print("\n拆分示例:")
    example_count = 0
    for idx, row in df.iterrows():
        if pd.notna(row['Phenotype(s)']) and ';' in str(row['Phenotype(s)']):
            print(f"\n原始 Phenotype: {row['Phenotype(s)']}")
            print(f"对应药物: {row['Drug(s)']}")
            phenotypes = [p.strip() for p in str(row['Phenotype(s)']).split(';')]
            print(f"拆分后: {phenotypes}")
            example_count += 1
            if example_count >= 3:  # 只显示3个示例
                break
    
    return new_df

def validate_preprocessing(original_file, processed_file):
    """
    验证预处理的正确性
    """
    print("\n" + "="*60)
    print("验证预处理结果...")
    
    original_df = pd.read_excel(original_file)
    processed_df = pd.read_excel(processed_file)
    
    # 检查列是否一致
    if list(original_df.columns) == list(processed_df.columns):
        print("✓ 列名保持一致")
    else:
        print("✗ 列名不一致")
        return False
    
    # 检查非Phenotype列的数据是否正确复制
    test_row = original_df[original_df['Phenotype(s)'].str.contains(';', na=False)].iloc[0] if any(original_df['Phenotype(s)'].str.contains(';', na=False)) else None
    
    if test_row is not None:
        original_phenotype = test_row['Phenotype(s)']
        phenotypes = [p.strip() for p in original_phenotype.split(';')]
        
        # 检查拆分后的行
        for phenotype in phenotypes:
            matched_rows = processed_df[
                (processed_df['Phenotype(s)'] == phenotype) & 
                (processed_df['Drug(s)'] == test_row['Drug(s)']) &
                (processed_df['Gene'] == test_row['Gene'])
            ]
            
            if len(matched_rows) > 0:
                print(f"✓ Phenotype '{phenotype}' 正确拆分并保留了原始数据")
            else:
                print(f"✗ Phenotype '{phenotype}' 拆分可能有问题")
    
    print("="*60)
    return True

def main():
    # 文件路径
    input_file = "/Users/qiaopan/Downloads/leo-matched_snp_genotypes_results666_reordered 2.xlsx"
    output_file = "/Users/qiaopan/Library/Mobile Documents/com~apple~CloudDocs/Work-Ebovir/全基因组测序项目/pharmacy/preprocessed_data.xlsx"
    
    try:
        # 预处理数据
        new_df = preprocess_phenotypes(input_file, output_file)
        
        # 验证结果
        validate_preprocessing(input_file, output_file)
        
        print("\n✅ 预处理完成！")
        print(f"下一步：使用处理后的文件 '{output_file}' 生成Word文档")
        
        return output_file
        
    except Exception as e:
        print(f"\n❌ 预处理失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    main()
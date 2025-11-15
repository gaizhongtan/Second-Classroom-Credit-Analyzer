import openpyxl
import pandas as pd
import warnings

# --- 初始化与用户输入 ---
# 禁止显示UserWarning
warnings.filterwarnings("ignore", category=UserWarning)

# 获取用户输入
name1 = input("请输入学院名称（不得简写）：\n")
总人数 = int(input("请输入学院总人数：\n"))
发布数量 = int(input("请输入学院活动发布数量：\n"))
被驳回活动总数 = int(input("请输入学院活动被驳回总数：\n"))
活动总数 = 发布数量
活动参与人次 = int(input("请输入学院活动参与人次：\n"))
活动签到率 = float(input("请输入学院活动签到率（例如输入85.5表示85.5%）：\n"))
学生总数 = 总人数
人均参与活动数 = float(input("请输入学院人均参与活动数：\n"))

# 根据学院名称定义相关文件名
details_file = f"{name1}活动明细.xlsx"
summary_file = f"{name1}学分汇总.xlsx"
template_file = "月报格式（学院） (2)(1).xlsx"  # 模板文件名
output_file = f"{name1}月报-已生成.xlsx"  # 最终生成的报告文件名

print("=" * 60)
print("第一部分：分析活动明细文件...")
print("=" * 60)

# --- 第一部分：分析《活动明细》 ---
# 读取活动明细Excel文件
try:
    df_details = pd.read_excel(details_file, engine='openpyxl')
except FileNotFoundError:
    print(f"错误：找不到文件 '{details_file}'。请确保文件名正确且文件存在于同一目录下。")
    exit()
except ValueError as e:
    print(f"读取 '{details_file}' 时出错: {e}")
    print("请确保该文件是有效的 .xlsx 文件。")
    exit()

# 1. 计算各类活动发放的总学分
total_credits_by_cat = df_details.groupby('活动分类')['发放学分总数'].sum()
print("各活动分类发放学分总数:")
print(total_credits_by_cat)
print('\n')

# 2. 计算补发学分
reissue_df = df_details[df_details['活动标题'].str.contains('补发', na=False)]
reissued_credits_by_cat = reissue_df.groupby('活动分类')['发放学分总数'].sum()
print('-----------以下是补发的-----------------')
if reissue_df.empty:
    print("未找到任何'补发'活动。")
else:
    print("含有'补发'的活动标题（前5条）:")
    print(reissue_df['活动标题'].head())
    print("\n各活动分类补发学分总数:")
    print(reissued_credits_by_cat)
print('------------------------------------')

# 3. 计算实际发放学分（总学分 - 补发学分）
actual_credits_by_cat = total_credits_by_cat.subtract(reissued_credits_by_cat, fill_value=0)
print('\n-----------实际发放分数（自动计算）-----------')
print(actual_credits_by_cat)
print('------------------------------------')

print("\n" + "=" * 60)
print("第二部分：分析学分汇总文件...")
print("=" * 60)

# --- 第二部分：分析《学分汇总》 ---
# 读取学分汇总Excel文件
try:
    df_summary = pd.read_excel(summary_file, engine='openpyxl')
except FileNotFoundError:
    print(f"错误：找不到文件 '{summary_file}'。请确保文件名正确且文件存在于同一目录下。")
    exit()
except ValueError as e:
    print(f"读取 '{summary_file}' 时出错: {e}")
    print("请确保该文件是有效的 .xlsx 文件。")
    exit()

# 1. 计算各项积分总和
积分_技能特长 = df_summary['技能特长积分-积分'].sum()
积分_创新创业 = df_summary['创新创业积分-积分'].sum()
积分_志愿公益 = df_summary['志愿公益积分-积分'].sum()
积分_工作履历 = df_summary['工作履历积分-积分'].sum()
积分_文体活动 = df_summary['文体活动积分-积分'].sum()
积分_思想成长 = df_summary['思想成长积分-积分'].sum()
积分_实践实习 = df_summary['实践实习积分-积分'].sum()
积分_总和 = df_summary['积分总和'].sum()

# 2. 计算人均分
if 总人数 > 0:
    人均分 = 积分_总和 / 总人数
else:
    人均分 = 0

# 存储各项积分总和到一个字典中
summary_totals = {
    '技能特长': 积分_技能特长, '创新创业': 积分_创新创业, '志愿公益': 积分_志愿公益,
    '工作履历': 积分_工作履历, '文体活动': 积分_文体活动, '思想成长': 积分_思想成长,
    '实践实习': 积分_实践实习, '总和': 积分_总和
}

print("各学院积分获取情况:")
for key, value in summary_totals.items():
    print(f"{key}积分-积分总和: {value}")
print(f"总人数: {总人数}")
print(f"人均: {人均分:.2f}")

print("\n" + "=" * 60)
print("第三部分：寻找二课之星...")
print("=" * 60)

# --- 第三部分：寻找“二课之星” ---
# 1. 寻找各类单项积分最高者
columns_to_check = ["技能特长积分-积分", "创新创业积分-积分", "志愿公益积分-积分", "工作履历积分-积分",
                    "文体活动积分-积分", "思想成长积分-积分", "实践实习积分-积分"]
result_data = []
for col in columns_to_check:
    max_value = df_summary[col].max()
    clean_col_name = col.replace('积分-积分', '')
    if pd.notna(max_value) and max_value > 0:
        max_names = df_summary[df_summary[col] == max_value]["姓名"].tolist()
        names_str = ", ".join(max_names)
        result_data.append([clean_col_name, names_str, max_value])
    else:
        result_data.append([clean_col_name, "无", 0])

stars_df = pd.DataFrame(result_data, columns=["活动类型", "姓名", "积分"])
print("--- 单项积分最高者 ---")
print(stars_df)

# 2. 寻找总分最高的“二课之星”
max_total_score = df_summary['积分总和'].max()
total_star_names = "无"
if pd.notna(max_total_score) and max_total_score > 0:
    total_star_names_list = df_summary[df_summary['积分总和'] == max_total_score]['姓名'].tolist()
    total_star_names = ", ".join(total_star_names_list)

print("\n--- 总分最高二课之星 ---")
print(f"姓名: {total_star_names}, 最高总分: {max_total_score if pd.notna(max_total_score) else 0}")

print("\n" + "=" * 60)
print("第四部分：生成Excel报告文件...")
print("=" * 60)

# --- 第四部分：将结果写入Excel模板 ---
try:
    wb = openpyxl.load_workbook(template_file)
    sheet = wb.active

    # (一) 各学院活动发布数量及驳回情况
    sheet['A4'] = name1
    sheet['B4'] = 发布数量
    sheet['C4'] = 被驳回活动总数

    # (二) 各学院参与人次
    sheet['A9'] = name1
    sheet['B9'] = 活动总数
    sheet['C9'] = 活动参与人次
    sheet['D9'] = f"{活动签到率}%"
    sheet['F9'] = 学生总数

    # (三) 各学院积分发放情况 -- 【已修正此处的总和计算逻辑】
    CREDITS_DIST_MAP_NO_SUM = {'技能特长': 'B', '创新创业': 'C', '志愿公益': 'D', '工作履历': 'E', '文体活动': 'F',
                               '思想成长': 'G', '实践实习': 'H'}
    sheet['A15'] = name1

    # 初始化用于累加的变量
    total_sum_15, total_sum_16, total_sum_17 = 0, 0, 0

    # 遍历所有分类，写入数据并累加
    for cat, col_letter in CREDITS_DIST_MAP_NO_SUM.items():
        val15 = total_credits_by_cat.get(cat, 0)
        val16 = reissued_credits_by_cat.get(cat, 0)
        val17 = actual_credits_by_cat.get(cat, 0)

        sheet[f'{col_letter}15'] = val15
        sheet[f'{col_letter}16'] = val16
        sheet[f'{col_letter}17'] = val17

        total_sum_15 += val15
        total_sum_16 += val16
        total_sum_17 += val17

    # 将计算好的总和写入I列
    sheet['I15'] = total_sum_15
    sheet['I16'] = total_sum_16
    sheet['I17'] = total_sum_17

    # (四) 获得积分
    CREDITS_DIST_MAP_WITH_SUM = {'技能特长': 'B', '创新创业': 'C', '志愿公益': 'D', '工作履历': 'E', '文体活动': 'F',
                                 '思想成长': 'G', '实践实习': 'H', '总和': 'I'}
    sheet['A26'] = name1
    for cat, col_letter in CREDITS_DIST_MAP_WITH_SUM.items():
        sheet[f'{col_letter}26'] = summary_totals.get(cat, 0)
    sheet['J26'] = 人均分
    sheet['K26'] = 总人数

    # (五) 学院人均参与活动数
    sheet['A32'] = name1
    sheet['B32'] = 人均参与活动数

    # 个人情况 (一) 各类积分获得数量最高 (单项)
    STARS_ROW_MAP = {'思想成长': 39, '创新创业': 40, '技能特长': 41, '志愿公益': 42, '文体活动': 43, '实践实习': 44,
                     '工作履历': 45}
    for index, row in stars_df.iterrows():
        cat = row['活动类型']
        if cat in STARS_ROW_MAP:
            target_row = STARS_ROW_MAP[cat]
            sheet[f'B{target_row}'] = row['姓名']
            sheet[f'C{target_row}'] = row['积分']

    # 个人情况 (一) 二课之星 (总分)
    sheet['A49'] = name1
    sheet['B49'] = total_star_names
    sheet['C49'] = max_total_score if pd.notna(max_total_score) and max_total_score > 0 else 0

    # （二）工作数据
    sheet['A53'] = name1
    sheet['B53'] = 活动总数
    sheet['C53'] = 被驳回活动总数
    sheet['D53'] = 活动参与人次
    sheet['E53'] = f"{活动签到率}%"
    sheet['F53'] = total_sum_17
    # 保存为新文件
    wb.save(output_file)
    print(f"成功！已将所有结果填入模板并保存为新文件：'{output_file}'")

except FileNotFoundError:
    print(f"错误：找不到模板文件 '{template_file}'。请确保模板文件存在于同一目录下。")
except Exception as e:
    print(f"写入Excel时发生错误: {e}")
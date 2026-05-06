# H&M Personalized Fashion Recommendations

基于 Kaggle H&M 竞赛的个性化时尚推荐系统，通过"时间衰减 + 位置加权复购"规则模型实现 MAP@12 = 0.02229（+19% vs Baseline）。

## 核心发现

- **复购是第一驱动力**：40%+ 用户在 31 天内有重复购买行为，复购信号的预测能力远超用户画像
- **近期商品热度是第二强信号**：从众效应使得"大家都在买"本身就是推荐依据
- **用户年龄等人口统计特征贡献接近于零**：经消融实验与 SHAP 分析验证，属于"伪重要性"特征

## 模型对比

| 版本 | MAP@12 | 关键改进 |
|------|--------|----------|
| Baseline（全局热门） | 0.01873 | 无个性化 |
| V12（时间衰减） | 0.02085 | `exp(-0.1 × days)` |
| **V13（位置加权复购）** | **0.02229** | 位置权重 + 时间衰减 + 冷启动补位 |

## 文件清单

| 文件 | 说明 |
|------|------|
| `期末报告.md` | 完整数据分析报告（Executive Summary、EDA、特征工程、消融实验、失败分析、建议） |
| `myCode.ipynb` | V13 最终模型代码（数据加载、EDA、加权复购模型、提交生成） |
| `EDA_Checkpoint.ipynb` | 数据探索与清洗（缺失值填充、TF-IDF、可视化） |
| `myNotebook.ipynb` | LightGBM 消融实验与 SHAP 可解释性分析 |
| `notebook.ipynb` | 关联规则基线流水线 |
| `H&M_个性化推荐系统_数据分析报告.pptx` | 演示文稿 |
| `H&M个性化推荐实验报告.md` | 实验报告（期中版本） |
| `漂移现象分析.md` | 数据漂移与概念漂移分析 |

## 环境

- Python 3.10.9 · Kaggle Notebook
- `lightgbm==4.3.0` · `pandas==2.2.2` · `numpy==1.23.5` · `scikit-learn==1.5.1`

## 竞赛链接

https://www.kaggle.com/competitions/h-and-m-personalized-fashion-recommendations

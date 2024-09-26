import pandas as pd
from matplotlib import pyplot as plt
from sklearn.cluster import KMeans
from scipy.spatial import distance

# 读取Excel数据
df = pd.read_excel(r'C:\Users\kd35\Desktop\井口坐标.xlsx')
# 提取坐标数据
X = df[['横轴坐标x(m)', '纵轴坐标y(m)']]

# 使用KMeans算法聚类
kmeans = KMeans(n_clusters=2)  # 假设要聚成2个集气站
kmeans.fit(X)

# 聚类结果
df['Cluster'] = kmeans.labels_

# 计算聚类中心（联合站）的坐标
cluster_centers = kmeans.cluster_centers_
cluster_df = pd.DataFrame(cluster_centers, columns=['横轴坐标x(m)', '纵轴坐标y(m)'])
cluster_df.index.name = '联合站编号'

# 计算每个联合站到各个井口的距离
distances = {}
for i, center in enumerate(cluster_centers):
    distances[f'联合站{i+1}'] = [distance.euclidean(center, point) for point in X.values]

# 构建DataFrame保存距离数据
distance_matrix = pd.DataFrame(distances, index=df['井号'])

# 保存结果到新的Excel表
with pd.ExcelWriter(r'C:\Users\kd35\Desktop\聚类结果.xlsx') as writer:
    df.to_excel(writer, sheet_name='原始数据', index=False)
    cluster_df.to_excel(writer, sheet_name='联合站坐标')
    distance_matrix.to_excel(writer, sheet_name='距离矩阵')

    # 绘制散点图
    plt.figure(figsize=(10, 8))
    plt.rcParams['font.sans-serif'] = ['SimHei']  # 指定默认字体为 SimHei
    plt.rcParams['axes.unicode_minus'] = False  # 解决保存图像是负号'-'显示为方块的问题

    # 绘制井口
    plt.scatter(df['横轴坐标x(m)'], df['纵轴坐标y(m)'], c=df['Cluster'], cmap='viridis', label='井口')
    # 绘制联合站（聚类中心）
    plt.scatter(cluster_centers[:, 0], cluster_centers[:, 1], marker='*', s=200, c='red', label='联合站')
    first_station =(6631.72,8435.17)
    plt.scatter(first_station[0], first_station[1], marker='s', s=200, c='blue', label='首站')

    # 添加标签和标题
    plt.xlabel('横轴坐标x(m)')
    plt.ylabel('纵轴坐标y(m)')
    plt.title('井口和联合站分布图')

    # 添加图例
    plt.legend()

    # 显示图形
    plt.grid(True)
    plt.show()
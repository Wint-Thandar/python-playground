from sklearn.cluster import KMeans
import pandas as pd
import matplotlib.pyplot as plt


# import sample_data_clustering.csv file from data folder as dataframe
df = pd.read_csv("python-playground/unsupervised-k-means-clustering/data/sample_data_clustering.csv")


# plot the data
plt.scatter(df['var1'], df['var2'])
plt.xlabel("var1")
plt.ylabel("var2")
plt.show()


# instantiate & fit the model
kmeans = KMeans(n_clusters=3, random_state=42)
kmeans.fit(df)


# add the cluster lables to our df
df['cluster'] = kmeans.labels_
df['cluster'].value_counts()


# plot our clusters and centroids
centroids = kmeans.cluster_centers_

clusters = df.groupby('cluster')

for cluster, data in clusters:
    plt.scatter(data['var1'], data['var2'], marker='o', label=cluster)
    plt.scatter(centroids[cluster, 0], centroids[cluster, 1], color='black', marker='X', s=300)
plt.legend()
plt.tight_layout()
plt.show()

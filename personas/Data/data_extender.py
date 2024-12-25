import numpy as np
import pandas as pd

# Sample means and standard deviations (calculated from the original dataset)
# For illustrative purposes, using arbitrary values here
means_class_0 = np.array([np.log(50000)] * 20)
stds_class_0 = np.array([1.0] * 20)

means_class_1 = np.array([np.log(10000)] * 20)
stds_class_1 = np.array([1.5] * 20)

# Number of samples per class
n_class_0 = 150
n_class_1 = 100

# Generate data for Class 0
data_class_0 = np.random.normal(loc=means_class_0, scale=stds_class_0, size=(n_class_0, 20))
data_class_0 = np.exp(data_class_0)
labels_class_0 = np.zeros((n_class_0, 1))

# Generate data for Class 1
data_class_1 = np.random.normal(loc=means_class_1, scale=stds_class_1, size=(n_class_1, 20))
data_class_1 = np.exp(data_class_1)
labels_class_1 = np.ones((n_class_1, 1))

# Combine the data
data = np.vstack((data_class_0, data_class_1))
labels = np.vstack((labels_class_0, labels_class_1))

# Create a DataFrame
df = pd.DataFrame(data, columns=[f'Feature{i}' for i in range(1, 21)])
df['Classification'] = labels.astype(int)

# Shuffle the DataFrame
df = df.sample(frac=1).reset_index(drop=True)

# Save to CSV
df.to_csv('extended_data.csv', index=False)
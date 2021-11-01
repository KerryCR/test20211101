import numpy as np
import matplotlib.pyplot as plt

labels = 'Regular flights', 'Maneuvering flight'
fraces = [12059, 1895]
# explode = [0, 0.1, 0, 0]
plt.axes(aspect=1)
plt.pie(x=fraces, labels=labels, autopct='%0f%%', shadow=True)
plt.show()

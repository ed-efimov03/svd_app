import numpy as np
from skimage import io
import os

# === Настройки ===
filename = "Cat.jpg"

# === Загружаем изображение ===
img = io.imread(filename)

# Получаем имя файла без расширения
basename = os.path.splitext(os.path.basename(filename))[0]

output_dir = f"compressed_images_{basename}"  # папка для сохранения результатов

# Создаём папку, если её нет
os.makedirs(output_dir, exist_ok=True)

# Список значений k
k_list = list(range(1, 11)) + list(range(20, 101, 10))

# Определяем количество каналов 
if img.ndim == 2:   # чёрно-белое изображение
    num_channels = 1
else:               # цветное изображение (RGB, RGBA и т.д.)
    num_channels = img.shape[2]


for k in k_list:
    new_img = []  # для каждого k — заново собираем изображение

    # SVD по каждому каналу RGB
    for i in range(3):
        channel = img if num_channels == 1 else img[:, :, i]
        U, S, Vt = np.linalg.svd(channel, full_matrices=False)

        # Берём только первые k компонент
        U_k = U[:, :k]
        S_k = np.diag(S[:k])
        Vt_k = Vt[:k, :]

        # Восстанавливаем канал
        img_compressed_channel = U_k @ S_k @ Vt_k
        new_img.append(img_compressed_channel)

    # Объединяем каналы
    img_compressed = np.stack(new_img, axis=2)
    img_compressed = np.clip(img_compressed, 0, 255)
    img_compressed_uint8 = img_compressed.astype(np.uint8)

    # === Сохраняем результат ===
    output_path = os.path.join(output_dir, f"{basename}_compressed_with_{k}_components.jpg")
    io.imsave(output_path, img_compressed_uint8)



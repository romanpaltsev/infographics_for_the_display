import os


text = "1920x1080"

resolution_split = list(map(int, text.split("x")))

print(resolution_split[0] / resolution_split[1])
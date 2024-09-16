data = "Rohan is not very smart. He is in fact, kind of stupid.".upper()
data = "".join([char for char in data if char not in [" ", ",", ".", "'", "-", '"', "+", "-"]])
print(data)

alphabet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']

count = 0
added_spaces = 0

for i in range(len(data)):
    count += 1
    if count % 5 == 0:
        data = data[:count + added_spaces] + " " + data[count + added_spaces:]
        added_spaces += 1



print(count)
print(data)
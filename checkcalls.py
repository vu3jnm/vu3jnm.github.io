calllists = []
filecontents = []
with open("calllist.txt", "r") as f:
    calllists = [f"<td>{i.strip()}</td>" for i in f.readlines()]


with open("searchcall.html", "r") as f:
    filecontents = f.read()

for i in calllists:
    if not i in filecontents:
        print(f"<tr>{i}<td> </td><td> </td> <td> </td> <td> </td></tr>")

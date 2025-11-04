import requests

calllists = []
filecontents = []
with open("calllist.txt", "r") as f:
    calllists = [f"<td>{i.strip()}</td>" for i in f.readlines()]


with open("searchcall.html", "r") as f:
    filecontents = f.read()

for i in calllists:
    if not i in filecontents:
        print(f"<tr>{i}<td> </td><td> </td> <td> </td> <td> </td></tr>")


url = "https://vu3jnm.co.in/searchcall.html"

content = requests.get(url).text

contents_list = content.split("\n")
print("in web")
for i in calllists:
    if not i in content:
        print(f"<tr>{i}<td> </td><td> </td> <td> </td> <td> </td></tr>")

print("[WEB]Find Dupplicates....")


for i in calllists:
    if content.count(i)>1:
        print(f"{i} duplicated")

from win32com.client import Dispatch

server = Dispatch("{AF080472-F173-4D9D-8BE7-435776617347}")

try:
    print('server.MethodWithArrayArgs(["param1"], ["param2"])')
    print(server.MethodWithArrayArgs(["param1"], ["param2"]))
except Exception as e:
    print(e)

print()

try:
    print('server.MethodWithArrayArgs(["param1"], None)')
    print(server.MethodWithArrayArgs(["param1"], None))
except Exception as e:
    print(e)

print()

try:
    print('server.Name = "MyName"')
    server.Name = "MyName"
    print(server.Name)
except Exception as e:
    print(e)

print()

try:
    print('server.Name = None')
    server.Name = None
    print(server.Name)
except Exception as e:
    print(e)

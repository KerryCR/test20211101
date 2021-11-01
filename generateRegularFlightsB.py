from typing import Tuple

from gurobipy import *
import xlrd
import xlwt

RegularFlightsA_xls = xlrd.open_workbook('RegularFlightsB.xls')  # xls
RegularFlightsATable = RegularFlightsA_xls.sheet_by_name('RegularFlightsB')  # Time-Ranked Table
RegularFlightsAList = []  # Time-Ranked List
for i in range(0, RegularFlightsATable.nrows):
    RegularFlightsAList += [RegularFlightsATable.row(i)]
# print(RegularFlightsAList) # RegularFlightsAList: Time-Ranked List
d = range(0, 30)  # !!!天数，跑B套时要改！
dNum = 30
M: Tuple[int, ...] = tuple([100000])  # 极大数
F = list(zip(*RegularFlightsAList))[0]  # 固定航班的航班号集合
Sd = list(zip(*RegularFlightsAList))[3]  # 固定航班的出发地集合
Sa = list(zip(*RegularFlightsAList))[6]  # 固定航班的到达地集合
Td = list(zip(*RegularFlightsAList))[8]  # 固定航班的出发时间点集合（以5min为单位）
Ta = list(zip(*RegularFlightsAList))[9]  # 固定航班的到达时间点集合（以5min为单位）
# print(type(Td), Td)
n = len(F)  # 固定航班个数
Pnn = [[0 for i in range(n)] for j in range(n)]  # 生成地点衔接矩阵Pnn，pij代表航班i的尾是否可接j的头
# print(type(Pnn), Pnn)
for i in range(0, n):
    for j in range(0, n):
        tmp1 = str(Sa[i])
        tmp2 = str(Sd[j])
        if tmp1 == tmp2:
            Pnn[i][j] = 1
# print(type(Pnn), Pnn)
Pdn = [0 for i in range(0, n)]  # 生成出发地标志列表Pdn，pdi表示航班i的出发地是否为基地TGD:Sd[0]或基地HOM:Sd[10]
for i in range(0, n):
    if str(Sd[i]) == str(Sd[0]):
        Pdn[i] = 1
Pdn = tuple(Pdn)
# print(type(Pdn), Pdn)
Pan = [0 for i in range(0, n)]  # 生成到达地标志列表Pan，pai表示航班i的到达地是否为基地TGD:Sd[0]或基地HOM:Sd[10]
for i in range(0, n):
    if str(Sa[i]) == str(Sd[0]):
        Pan[i] = 1
# print(type(Pan), Pan)
Tnn = [[0 for i in range(n)] for j in range(n)]  # 生成时间衔接矩阵Tnn，pij代表航班i的尾是否可接j的头
for i in range(0, n):
    for j in range(0, n):
        if int(Td[j].value) - int(Ta[i].value) >= 8:
            Tnn[i][j] = 1
# print(type(Tnn), Tnn)

m = Model("generateRegularFlights")
NL = 40  # !!!固定航线数目，需要手动调整至收敛！
l = m.addVars(n, NL, vtype=GRB.BINARY, name="l")  # 添加n*NL的变量，命名为l
FightUse = m.addVars(n, 1, vtype=GRB.BINARY, name="FightUse")  # 添加n*1的变量，命名为FightUse 存放班机是否使用的变量
temp_n = range(0, n)
temp_NL = range(0, NL)
# expr = sum(l[i,j] for i in temp_n for j in temp_NL)
# expr = sum((0 if (sum(l[i, j] for j in temp_NL) == 0) else 1) for i in temp_n)
# expr = sum(1 if (sum(l[i, j] for j in temp_NL)) == 0 else 0 for i in temp_n)
expr = sum(FightUse[i, 0] for i in temp_n)  # FightUse变量存储的是每一行之和是否大于0，大于0则为1  要让每个班机尽可能使用
m.setObjective(expr, GRB.MAXIMIZE)  # 建立目标函数

# if else的逻辑约束转换  FightUse变量存储的是每一行之和是否大于0，大于0则为1
for i in temp_n:
    m.addConstr(sum(l[i, j] for j in temp_NL) - 10000 * FightUse[i, 0] <= 0, "FightUse{i}")
    m.addConstr(FightUse[i, 0] - 10000 * sum(l[i, j] for j in temp_NL) <= 0, "FightUse{i}")
    # m.addConstr(sum(l[i, j] for j in temp_NL) - 0.5 <= 10000 * FightUse[i, 0], "FightUse{i}")
    # m.addConstr(-sum(l[i, j] for j in temp_NL) + 0.5 <= 10000 * (1 - FightUse[i, 0]), "FightUse{i}")

# 出发地是基地的约束
m.addConstrs((l[0, i] - Pdn[0] * 10000 <= 0 for i in temp_NL), "based{i}{j}")
for i in temp_NL:
    for j in range(1, n):
        m.addConstr(l[j, i] - 2 * sum(l[k, i] for k in range(0, j)) <= 10000 * Pdn[j], "based{i}{j}")
# 到达地是基地的约束
m.addConstrs((l[n - 1, i] - Pan[n - 1] * 10000 <= 0 for i in temp_NL), "basea{i}{j}")
for i in temp_NL:
    for j in range(0, n - 1):
        m.addConstr(l[j, i] - 2 * sum(l[k, i] for k in range(j, n)) <= 10000 * Pdn[j], "basea{i}{j}")
# 航线内航班地点互相衔接的约束
for i in temp_NL:
    for j in temp_n:
        for k in range(j + 1, n):
            m.addConstr(l[j, i] + l[k, i] - 1 - 10000 * sum(l[m, i] for m in range(j + 1, k)) - 10000 * Pnn[j][k] <= 0,
                        "P{i}{j}")
# 航线内航班时间相互衔接的约束
for i in temp_NL:
    for j in temp_n:
        for k in range(j + 1, n):
            m.addConstr(l[j, i] + l[k, i] - 1 - 10000 * sum(l[m, i] for m in range(j + 1, k)) - 10000 * Tnn[j][k] <= 0,
                        "T{i}{j}")
m.optimize()
m.printStats()
x = m.getVars()  # 决策变量值
for it in x:
    print(it)
obj = m.getObjective()
print(obj.getValue())
m.write("outputFlightsB.sol")

from difflib import get_close_matches
import pandas as pd
import xlsxwriter

# Read Excel
payMerchantDF = pd.read_excel('./PayMerchantInput.xlsx')

# Setup List by Rows
payMerchantAddressList = payMerchantDF["주소"].values.tolist()
payMerchantAddressList.sort()

# 상호명: 0, 대표자: 1, 주소: 7
# Similar strings are defined by a similarity ratio of 0.6 or greater
groups = []
while payMerchantAddressList:
    currAdd = payMerchantAddressList[0]
    similarList = get_close_matches(currAdd, payMerchantAddressList, cutoff=0.8)

    groups.append(similarList)
    
    # Remove addresses in similar list
    for similarAdd in similarList:
        payMerchantAddressList.remove(similarAdd)

# Rebuild Entire Excel Data from Groups
fullAddressGroups = []
fullDataSingles = []
payMerchantMatrix = payMerchantDF.values.tolist()

for group in groups:
    group.sort()

    fullAddressGroup = []

    for address in group:
        for payMerchant in payMerchantMatrix:
            if (address == payMerchant[2]):
                fullAddressGroup.append(payMerchant)
                payMerchantMatrix.remove(payMerchant)
                break

    if len(fullAddressGroup) == 1:
        fullDataSingles.append(fullAddressGroup[0])
    else:
        fullAddressGroups.append(fullAddressGroup) 


# Filter by Business Names from Groups
fullDataGroups = []
for fullAddressGroup in fullAddressGroups:
    # Build Business Name List from Address Groups
    nameList = []
    for i in range(0, len(fullAddressGroup)):
        nameList.append(fullAddressGroup[i][0])

    similarNameGroups = []

    while nameList:
        currName = nameList[0]
        similarNameList = get_close_matches(currName, nameList)
        similarNameGroups.append(similarNameList)

        # Remove Similar Names from nameList Queue
        for similarName in similarNameList:
            nameList.remove(similarName)

    # Rebuild Data with similarNameGroups
    fullDataGroup = []
    for similarNameGroup in similarNameGroups:
        for similarName in similarNameGroup:
            for data in fullAddressGroup:
                if (data[0] == similarName):
                    fullDataGroup.append(data)
                    fullAddressGroup.remove(data)
                    break

    fullDataGroups.append(fullDataGroup)

finalGroups = []
finalSingles = []

for singles in fullDataSingles:
    finalSingles.append(singles)

for fullDataGroup in fullDataGroups:
    if len(fullDataGroup) == 1:
        finalSingles.append(fullDataGroup)
    else:
        finalGroups.append(fullDataGroup)

finalCount = len(finalSingles)
for group in finalGroups:
    finalCount += len(group)

print(len(payMerchantDF))
print(finalCount)

# Create Excel from Groups
groupOutput = []
singlesOutput = finalSingles

print(singlesOutput)

for group in finalGroups:
    for item in group:
        groupOutput.append(item)
    groupOutput.append(["", "", ""])

groupDF = pd.DataFrame(groupOutput, columns = ["가맹점 명", "대표자", "주소"])
singlesDF = pd.DataFrame(singlesOutput, columns = ["가맹점 명", "대표자", "주소"])

writer = pd.ExcelWriter('PayMerchantNameAddressGroup.xlsx')
groupDF.to_excel(writer, sheet_name='비슷한 가맹점 목록')
singlesDF.to_excel(writer, sheet_name='그 외 목록')

writer.save()

# # Rebuild Data using Groups
# 
# output = []

# for group in groups:
#     group.sort()
#     for address in group:
#         for payMerchant in payMerchantMatrix:
#             if (address == payMerchant[2]):
#                 output.append(payMerchant)
#                 payMerchantMatrix.remove(payMerchant)
#                 break

#     output.append(["", "", ""])

# outputDF = pd.DataFrame(output, columns = ["가맹점 명", "대표자", "주소"])
# outputDF.to_excel("PayMerchantOutput.xlsx", index=False)
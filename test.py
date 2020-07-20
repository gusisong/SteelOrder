import pandas as pd
import datetime

data = pd.read_excel('OrderInfo/Template.xlsx', index_col=0, header=1)
data.rename(columns={'料片\n供应商': '订货单位'}, inplace=True)

plant_list = list(set(data['工厂']))

for plant in plant_list:
    fil_plant = data[data.工厂 == plant]
    buyer_list = list(set(fil_plant['订货单位']))

    for buyer in buyer_list:
        fil_buyer = fil_plant[fil_plant.订货单位 == buyer]

        # fil_buyer.to_csv('Detail\{0}_{1}.csv'.format(plant, buyer), encoding='gb2312')

        part_number_list = list(fil_buyer['物料编号'])

        count = 0
        for i in range(0, len(part_number_list), 10):
            part_number_group = part_number_list[i:i + 10]

            n = 0
            for part_number in part_number_group:
                print(count + n, part_number)
                n += 1

            count += 10
            print('*' * 50)

        break
    break

print(datetime.datetime.now().month)

print(data.shape[0])
v_d_list = []

v_ch_list = []

with open(r"V:\Hyundai Mobility\Outgoing\Копии документов ТС\УБЫТКИ\Учет убытков\Сверка с ВСК\Выгрузка_убытков\all_damages.txt") as vins_damages:

    for line in vins_damages:
        line = line.upper().strip().split('\t')
        v_d_list.append(line)

    # print(v_d_list)
    # print()

    for el in v_d_list:
        el[1] = el[1].replace('"', '')

    # print(v_d_list)
    # print()

    vins_damages_list = []

    for pair in v_d_list:
        v_d_dict = {'VIN': pair[0], 'DAMAGES': pair[1]}
        vins_damages_list.append(v_d_dict)

    # print(vins_damages_list)
    # print()

    with open(r"V:\Hyundai Mobility\Outgoing\Копии документов ТС\УБЫТКИ\Учет убытков\Сверка с ВСК\Выгрузка_убытков\vins_for_checking.txt", encoding='utf8') as vins_for_checking:

        for line in vins_for_checking:
            line = line.upper().strip()
            v_ch_list.append(line)

        # print(v_ch_list)
        # print()

        v_ch_dict = {}
        for vin in v_ch_list:
            v_ch_dict[vin] = []

        # print(v_ch_dict)
        # print()

        for vin, damages in v_ch_dict.items():
            for dict_ in vins_damages_list:
                if vin == dict_['VIN']:
                    damages.append(dict_['DAMAGES'])

        for vin, damages in v_ch_dict.items():
            if not damages:
                damages.append('убытков_нет')

        # print(v_ch_dict)
        # print()

        with open(r"V:\Hyundai Mobility\Outgoing\Копии документов ТС\УБЫТКИ\Учет убытков\Сверка с ВСК\Выгрузка_убытков\input.txt", 'w', encoding='utf8') as input:

            for vin, damages in v_ch_dict.items():
                input.write(f"{vin}§{'§'.join(damages)} \n")

print(r"Убытки найдены. Откройте файл Все_убытки.xls и выполните третий макрос")


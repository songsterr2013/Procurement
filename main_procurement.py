from Procurement import Procurement
import traceback
import util
import os

logger = util.get_logger(__file__)

if __name__ == '__main__':

    s, config = util.load_config()

    if s:
        try:
            main_excel = config["main_excel"]
            alt_main_excel = config["alt_main_excel"]
            save_path = config["save_path"]
            bom_path = config["bom_path"]
            pallet_path = config["pallet_path"]

            excel = ''
            way_to_go = None

            # 如果整張都是拆單的話，棧板就只取一次就break，因為看似採購的那張EXCEL會跟炸製令一模一樣(lazy mode)
            if os.path.exists(main_excel):
                excel = main_excel
            elif os.path.exists(alt_main_excel):
                way_to_go = True
                excel = alt_main_excel

            p = Procurement(excel, save_path, bom_path, pallet_path)
            p.make_new_sheet()

            # 用遍歷每一個row的方式去處理
            for item in p.read_main_excel():

                tmp_pallet_list = []
                tmp_bom_list = []

                for i in p.get_required_pallet_data(item[2], item[6]):
                    tmp_pallet_list.append(i)
                    if way_to_go:
                        break

                for q in p.get_required_bom_data(item[6]):
                    tmp_bom_list.append(q)

                p.write_data(item, tmp_pallet_list, tmp_bom_list)

            p.save()

        except ValueError:
            logger.error("錯誤:{}".format(traceback.format_exc()))

    else:
        logger.error("沒有設定檔或設定檔異常：{}".format(traceback.format_exc()))

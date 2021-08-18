from Procurement import Procurement
import traceback
import util

logger = util.get_logger(__file__)

if __name__ == '__main__':

    s, config = util.load_config()

    if s:
        try:
            main_excel = config["main_excel"]
            parse_local = config["parse_local"]
            bom_path = config["bom_path"]
            pallet_path = config["pallet_path"]

            p = Procurement(main_excel, parse_local, bom_path, pallet_path)

            # 用遍歷每一個row的方式去處理
            for item in p.read_main_excel():
                p.get_required_data(item[6])
                break
        except ValueError as e:
            logger.error("錯誤:{}".format(traceback.format_exc()))

    else:
        logger.error("沒有設定檔或設定檔異常：{}".format(traceback.format_exc()))

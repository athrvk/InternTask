from util_xlwings import intern_test
import time

test_result = intern_test()


def run():
    test_result.create_xlfile('test.xlsx')
    wb = test_result.load_xlfile('test.xlsx')
    test_result.initialize_workbook(wb)

    print()

    while test_result.flag:
        r = test_result.req_data()
        time.sleep(0.2)
        test_result.append_data(r, wb)
        time.sleep(0.2)
        test_result.save_file(wb, 'test.xlsx')
        time.sleep(0.2)
        test_result.check_temp(wb)
        time.sleep(0.2)
        test_result.check_state(wb)
        time.sleep(0.2)
        print()


if __name__ == '__main__':
    run()

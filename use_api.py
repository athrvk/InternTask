from util_xlwings import intern_test
import time

test_result = intern_test()


def start_the_program():
    test_result.create_xlfile('test.xlsx')
    workbook = test_result.load_xlfile('test.xlsx')
    test_result.initialize_workbook(workbook)

    print()

    while not test_result.check_update_flag:
        r = test_result.request_data()
        time.sleep(0.2)

        test_result.append_data(r, workbook)
        time.sleep(0.2)

        test_result.save_file(workbook, 'test.xlsx')
        time.sleep(0.2)

        test_result.is_temperature_C_or_F(workbook)
        time.sleep(0.2)

        test_result.to_stop_updating(workbook)
        time.sleep(0.2)
        print()


if __name__ == '__main__':
    start_the_program()

from util_xlwings import intern_test
import time

test_result = intern_test()


def start_the_program():
    test_result.create_xlfile('test.xlsx')
    workbook = test_result.load_xlfile('test.xlsx')
    test_result.initialize_workbook(workbook)

    print()

    while not test_result.exit_code:
        while test_result.check_update_flag and (not test_result.exit_code):
            r = test_result.request_data()
            time.sleep(0.3)

            test_result.append_data(r, workbook)
            time.sleep(0.3)

            test_result.save_file(workbook, 'test.xlsx')
            time.sleep(0.3)

            test_result.is_temperature_C_or_F(workbook)
            time.sleep(0.3)

            test_result.status_check(workbook)
            time.sleep(1)
            print()

        test_result.status_check(workbook)
        time.sleep(2)

    # workbook.close()
    print("Exiting Program...")


if __name__ == '__main__':
    start_the_program()

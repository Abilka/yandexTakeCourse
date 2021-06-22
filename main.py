from selenium import webdriver
import xlwt
import threading
import mail


class Browser:
    def __init__(self):
        self.driver = webdriver.Chrome()

    def take_course_yandex(self, button_class_name: str, money_name: str) -> None:
        self.driver.get('https://yandex.ru/')

        self.driver.find_element_by_xpath(f'//*[@class="{button_class_name}"]/a').click()
        self.driver.switch_to.window(self.driver.window_handles[-1])
        course = self.driver.find_element_by_xpath('//*[@class="news-stock-table__content"]')

        value[money_name] = []

        for course_line in course.find_elements_by_tag_name('div'):
            course_line_info = course_line.find_elements_by_tag_name('div')

            if len(course_line_info) > 2:
                if course_line_info[0].text != 'Дата':
                    value[self.money_name].append({'date': course_line_info[0].text,
                                                   'course': course_line_info[1].text,
                                                   'change': course_line_info[2].text})


class Exel:
    def __init__(self):
        self.wb = xlwt.Workbook()
        self.ws = self.wb.add_sheet('Курсы валют')
        self.row = 0
        self.col = 0
        self.all_amount_row = 0

    def create_course(self):
        for cur_name in value:
            for string in value[cur_name]:
                style = xlwt.XFStyle()
                style.num_format_str = '[$R]#,##0.00'
                self.ws.write(self.row, self.col, string["date"], style)
                self.ws.write(self.row, self.col + 1, string["course"], style)
                self.ws.write(self.row, self.col + 2, string["change"], style)
                self.row += 1
                self.all_amount_row = self.row
            self.col += 3
            self.row = 0
        for row in range(1, 11):
            style = xlwt.XFStyle()
            style.num_format_str = '0.00'
            self.ws.write(row-1, self.col, xlwt.Formula(f'$B{row}/$E{row}'), style)
        self.wb.save('Course.xls')



if __name__ == '__main__':
    value = {}
    threading.Thread(target=Browser().take_course_yandex, args=('b-inline inline-stocks__item inline-stocks__item_id_2002 hint__item inline-stocks__part', 'dollar')).start()
    Browser().take_course_yandex('b-inline inline-stocks__item inline-stocks__item_id_2000 hint__item inline-stocks__part','euro')

    courseFile = Exel()
    courseFile.create_course()
    mail.mail_send("Course.xls", courseFile.all_amount_row)

import mpmath
import openpyxl
import matplotlib.pyplot as plt
import pandas as pd

class Tether:
    def __init__(self):
        pass

    def weight_of_system(self, length, diametr, mka):
        """
        This function calculating weight of system of spacecraft and tether system
        :param length: length of tether system
        :param diametr: diametr of tether system
        :param mka: weigth of spacecraft
        :return: weight of system of spacecraft and tether system
        """
        self.weight_tros = 7660 * mpmath.pi * diametr * diametr * length / 4
        self.weight_system = mka + self.weight_tros
        return (self.weight_system)

    def coordinates_of_spacecraft(self, perigee, Radius_Earth=6371000):
        """
        This function calculating coordinates of spacecraft
        :param perigee: Height of the orbit
        :param Radius_Earth: Radius of Earth
        :return: coordinate "X" and "Y" of the spacecraft
        """
        self.r = Radius_Earth + perigee
        self.x = 0.5 * self.r
        self.y = mpmath.sqrt(self.r * self.r - self.x * self.x)
        return (self.x, self.y)

    def magnetic_induction(self, x, y, Kepler_i, mu0=1.25663706212 * pow(10, -6), Pm=8.3 * pow(10, 22)):
        """
        This function calculate tilt of the earth's magnetic field axis and return value of function (here cos()) depending on tilt of the earth's magnetic field axis.
        :param x: coordinate "X" of the spacecraft
        :param y: coordintae "Y" of the spacecraft
        :param Kepler_i: orbital inclination
        :param Pm: value of magnetic dipole moment
        :return: value of function (here cos()) depending on tilt of the earth's magnetic field axis.
        """
        self.r_square = self.x * self.x + self.y * self.y
        self.bz = -mu0 * (1 / (4 * mpmath.pi)) * (Pm / (self.r_square * mpmath.sqrt(self.r_square)))
        self.Btz = self.bz
        self.cosinus = (self.Btz * self.bz) / (self.Btz * self.bz)
        self.coskvim = (
                    6 + 2 * mpmath.cos(2 * Kepler_i) + 3 * mpmath.cos(2 * Kepler_i) + 2 * self.cosinus + 3 * mpmath.cos(
                2 * Kepler_i))
        return (self.coskvim)

    def deorbiting_time(self, weight_system, length, coskvim, perigee, apogee, heigth_atmosphere=120000,
                     Bm=4 * pow(10, -5), R=125 * pow(10, -8)):
        """
        This function returns time of deorbiting the spacecraft by using tether system
        :param weight_system: weight of system of spacecraft and tether system
        :param length: length of tether system
        :param coskvim: value of function (here cos()) depending on tilt of the earth's magnetic field axis.
        :param perigee: Height which from we deorbiting our spacecraft
        :param heigth_atmosphere: Height of dense layers of the atmosphere
        :param Bm: Magnetic induction at the geomagnetic equator
        :param R: Wire rope resistance
        :return: Time of deorbiting the spacecraft
        """
        self.Kepler_a = (apogee + perigee) / 2
        self.deltat = ((weight_system * R) / (2 * length * pow(self.Kepler_a, 6) * 1 * coskvim * Bm * Bm)) * (
                    pow(perigee, 7) - pow(heigth_atmosphere, 7))
        self.deltatmes = self.deltat * 3.8052 * pow(10, -7)
        print(
            f"Время увода системы массой {weight_system} с высоты орбиты {perigee} тросовой системой длинной {length} составляет {self.deltatmes} мес.")
        return self.deltatmes

    def write_data_into_excel(self, lst: list):
        """
        This function adding data with results to Excel sheets
        :param lst: 2 dimension list with results
        :return: Excel file with inserted data
        """
        sheets = ['L-5 000 M-900', 'L-10 000 M-900', 'L-15 000 M-900', 'L-20 000 M-900', 'L-25 000 M-900',
                  'L-5 000 M-1 800', 'L-10 000 M-1 800', 'L-15 000 M-1 800', 'L-20 000 M-1 800', 'L-25 000 M-1 800']
        title = ["Высота орбиты, м", "Масса КА, кг", "Масса системы, кг", "Длина троса, м", "Время увода, мес"]
        book = openpyxl.Workbook()
        for sheet in sheets:
            book.create_sheet(sheet)
            book[sheet].append(title)
        for i in lst:
            if i[1] == 900 and i[
                3] == 5000:  # or i == ["Высота орбиты, м", "Масса КА, кг", "Масса системы, кг", "Длина троса, м", "Время увода, мес"]:
                book['L-5 000 M-900'].append(i)
            elif i[1] == 900 and i[
                3] == 10000:  # or i == ["Высота орбиты, м", "Масса КА, кг", "Масса системы, кг", "Длина троса, м", "Время увода, мес"]:
                book['L-10 000 M-900'].append(i)
            elif i[1] == 900 and i[
                3] == 15000:  # or i == ["Высота орбиты, м", "Масса КА, кг", "Масса системы, кг", "Длина троса, м", "Время увода, мес"]:
                book['L-15 000 M-900'].append(i)
            elif i[1] == 900 and i[
                3] == 20000:  # or i == ["Высота орбиты, м", "Масса КА, кг", "Масса системы, кг", "Длина троса, м", "Время увода, мес"]:
                book['L-20 000 M-900'].append(i)
            elif i[1] == 900 and i[
                3] == 25000:  # or i == ["Высота орбиты, м", "Масса КА, кг", "Масса системы, кг", "Длина троса, м", "Время увода, мес"]:
                book['L-25 000 M-900'].append(i)
            elif i[1] == 1800 and i[
                3] == 5000:  # or i == ["Высота орбиты, м", "Масса КА, кг", "Масса системы, кг", "Длина троса, м", "Время увода, мес"]:
                book['L-5 000 M-1 800'].append(i)
            elif i[1] == 1800 and i[
                3] == 10000:  # or i == ["Высота орбиты, м", "Масса КА, кг", "Масса системы, кг", "Длина троса, м", "Время увода, мес"]:
                book['L-10 000 M-1 800'].append(i)
            elif i[1] == 1800 and i[
                3] == 15000:  # or i == ["Высота орбиты, м", "Масса КА, кг", "Масса системы, кг", "Длина троса, м", "Время увода, мес"]:
                book['L-15 000 M-1 800'].append(i)
            elif i[1] == 1800 and i[
                3] == 20000:  # or i == ["Высота орбиты, м", "Масса КА, кг", "Масса системы, кг", "Длина троса, м", "Время увода, мес"]:
                book['L-20 000 M-1 800'].append(i)
            elif i[1] == 1800 and i[
                3] == 25000:  # or i == ["Высота орбиты, м", "Масса КА, кг", "Масса системы, кг", "Длина троса, м", "Время увода, мес"]:
                book['L-25 000 M-1 800'].append(i)
        book.save('tether2.xlsx')

    def creating_a_diagram(self):
        """
        This function create plots and inserting it into Excel File with results
        :return: Excel file with diagrams
        """
        sheets = ['L-5 000 M-900', 'L-10 000 M-900', 'L-15 000 M-900', 'L-20 000 M-900', 'L-25 000 M-900',
                  'L-5 000 M-1 800', 'L-10 000 M-1 800', 'L-15 000 M-1 800', 'L-20 000 M-1 800', 'L-25 000 M-1 800']
        wb = openpyxl.load_workbook('tether2.xlsx')
        for sheet in sheets:
            ws = wb[sheet]
            var = pd.read_excel("tether2.xlsx", sheet_name=sheet)
            x = list(var['Время увода, мес'])
            y = list(var['Высота орбиты, м'])
            plt.figure(figsize=(7, 5))
            # plt.style.use('seaborn')
            plt.scatter(x, y, marker=".", s=100, edgecolors="black", c="yellow")
            plt.title("L = 5000m ")
            plt.xlabel('Deorbiting time', fontweight='bold', color='black', fontsize='12', horizontalalignment='center')
            plt.ylabel('Orbit`s height', fontweight='bold', color='black', fontsize='12', horizontalalignment='center')
            plt.plot(x, y, '-o')
            # plt.show()
            plt.savefig(f'{sheet}.png')
            img = openpyxl.drawing.image.Image(f'{sheet}.png')
            img.anchor = 'G1'
            ws.add_image(img)
            wb.save('tether2.xlsx')

    def calculate_all_results(self):
        """
        This function calculating all parameters, inserting it in Excel file, creating diagrams and inserting it in Excel File.
        :return: Excel file with calculated parameters and diagrams.
        """
        ######################################################
        #Initial data, which u can change
        ######################################################
        list_of_weights = [900, 1800]
        list_of_perigee = [700000, 800000, 900000]
        list_of_apogee = [800000, 900000, 1000000]
        length_of_tether = [5000, 10000, 15000, 20000, 25000]
        zipped_heights = list(zip(list_of_apogee, list_of_perigee))
        # print(zipped_heights)
        # create_file()
        lst = []
        results = [
            ["Высота орбиты, м", "Масса КА, кг", "Масса системы, кг", "Длина троса, м", "Время увода, мес"],
        ]
        for length in length_of_tether:
            for mass in list_of_weights:
                for tuple_of_heights in zipped_heights:
                    weigth = self.weight_of_system(length, 0.007, mass)
                    coordinates = self.coordinates_of_spacecraft(tuple_of_heights[0])
                    induction = self.magnetic_induction(coordinates[0], coordinates[1], 60)
                    time = float(self.deorbiting_time(weigth, length, induction, tuple_of_heights[0], tuple_of_heights[1]))
                    results.append([tuple_of_heights[0], int(mass), int(weigth), length, time])
        self.write_data_into_excel(results)
        self.creating_a_diagram()
        print(f'Result calculated!')


def main():
    tether = Tether()
    tether.calculate_all_results()

if __name__ == '__main__':
    main()


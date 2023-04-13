import mpmath
import openpyxl
import matplotlib.pyplot as plt
import pandas as pd

class Tros:
    def __init__(self):
        pass

    def weight(self, length, diametr, mka):
        """
        This function calculating weight of system of spacecraft and tross system
        :param length: length of tross system
        :param diametr: diametr of tross system
        :param mka: weigth of spacecraft
        :return: weight of system of spacecraft and tross system
        """
        self.weight_tros = 7660 * mpmath.pi * diametr * diametr * length / 4
        self.weight_system = mka + self.weight_tros
        return (self.weight_system)

    def coordinates(self, perigee, Radius_Earth=6371000):
        """
        This function calculating coordinates of spacecraft
        :param perigee: Height which from we deorbiting our spacecraft
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

    def deorbit_time(self, weight_system, length, coskvim, perigee, apogee, heigth_atmosphere=120000,
                     Bm=4 * pow(10, -5), R=125 * pow(10, -8)):
        """
        This function returns time of deorbiting the spacecraft by using tross system
        :param weight_system: weight of system of spacecraft and tros system
        :param length: length of tross system
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


"""def create_file():
    file_name = "./test.xlsx"
    book = openpyxl.Workbook('tether2.xlsx')
    book.create_sheet(' L-5 000 M-900')
    book.create_sheet(' L-10 000 M-900')
    book.create_sheet(' L-15 000 M-900')
    book.create_sheet(' L-20 000 M-900')
    book.create_sheet(' L-25 000 M-900')
    book.create_sheet(' L-5 000 M-1 800')
    book.create_sheet(' L-10 000 M-1 800')
    book.create_sheet(' L-15 000 M-1 800')
    book.create_sheet(' L-20 000 M-1 800')
    book.create_sheet(' L-25 000 M-1 800')
    book.save('tether2.xlsx')"""


def write_data_into_excel(lst: list):
    book = openpyxl.Workbook()
    book.create_sheet('L-5 000 M-900')
    book.create_sheet('L-10 000 M-900')
    book.create_sheet('L-15 000 M-900')
    book.create_sheet('L-20 000 M-900')
    book.create_sheet('L-25 000 M-900')
    book.create_sheet('L-5 000 M-1 800')
    book.create_sheet('L-10 000 M-1 800')
    book.create_sheet('L-15 000 M-1 800')
    book.create_sheet('L-20 000 M-1 800')
    book.create_sheet('L-25 000 M-1 800')
    title = ["Высота орбиты, м", "Масса КА, кг", "Масса системы, кг", "Длина троса, м", "Время увода, мес"]
    book['L-5 000 M-900'].append(title)
    book['L-10 000 M-900'].append(title)
    book['L-15 000 M-900'].append(title)
    book['L-20 000 M-900'].append(title)
    book['L-25 000 M-900'].append(title)
    book['L-5 000 M-1 800'].append(title)
    book['L-10 000 M-1 800'].append(title)
    book['L-15 000 M-1 800'].append(title)
    book['L-20 000 M-1 800'].append(title)
    book['L-25 000 M-1 800'].append(title)
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


def main():
    list_of_weights = [900, 1800]
    list_of_perigee = [700000, 800000, 900000]
    list_of_apogee = [800000, 900000, 1000000]
    length_of_tether = [5000, 10000, 15000, 20000, 25000]
    zipped_heights = list(zip(list_of_apogee, list_of_perigee))
    # print(zipped_heights)
    # create_file()
    lst = []
    rows = [
        ["Высота орбиты, м", "Масса КА, кг", "Масса системы, кг", "Длина троса, м", "Время увода, мес"],
    ]
    for length in length_of_tether:
        for mass in list_of_weights:
            for tuple_of_heights in zipped_heights:
                tros = Tros()
                weigth = tros.weight(length, 0.007, mass)
                coordinates = tros.coordinates(tuple_of_heights[0])
                induction = tros.magnetic_induction(coordinates[0], coordinates[1], 60)
                time = float(tros.deorbit_time(weigth, length, induction, tuple_of_heights[0], tuple_of_heights[1]))
                rows.append([tuple_of_heights[0], int(mass), int(weigth), length, time])
    print(rows)
    write_data_into_excel(rows)

def diagram():
    var = pd.read_excel("tether2.xlsx", sheet_name='L-5 000 M-900')
    x = list(var['Время увода, мес'])
    y = list(var['Высота орбиты, м'])
    plt.figure(figsize=(10, 10))
    #plt.style.use('seaborn')
    plt.scatter(x, y, marker=".", s=100, edgecolors="black", c="yellow")
    plt.title("L = 5000m ")
    plt.xlabel('Deorbiting time', fontweight='bold', color = 'black', fontsize='12', horizontalalignment='center')
    plt.ylabel('Orbit`s height', fontweight='bold', color = 'black', fontsize='12', horizontalalignment='center')
    plt.plot(x,y, '-o')
    plt.show()



if __name__ == '__main__':
    #main()
    #create_file()
    diagram()

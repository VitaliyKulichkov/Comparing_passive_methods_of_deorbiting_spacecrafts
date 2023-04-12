import mpmath
import openpyxl

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
        return(self.weight_system)

    def coordinates (self, perigee, Radius_Earth = 6371000):
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

    def magnetic_induction (self, x, y, Kepler_i, mu0 = 1.25663706212 *pow(10, -6), Pm = 8.3 * pow(10, 22) ):
        """
        This function calculate tilt of the earth's magnetic field axis and return value of function (here cos()) depending on tilt of the earth's magnetic field axis.
        :param x: coordinate "X" of the spacecraft
        :param y: coordintae "Y" of the spacecraft
        :param Kepler_i: orbital inclination
        :param Pm: value of magnetic dipole moment
        :return: value of function (here cos()) depending on tilt of the earth's magnetic field axis.
        """
        self.r_square = self.x*self.x + self.y*self.y
        self.bz = -mu0*(1/(4*mpmath.pi))*(Pm/(self.r_square*mpmath.sqrt(self.r_square)))
        self. Btz = self.bz
        self.cosinus = (self.Btz * self.bz) / (self.Btz * self.bz)
        self.coskvim = (6 + 2 * mpmath.cos(2 * Kepler_i) + 3 * mpmath.cos(2 * Kepler_i) + 2 * self.cosinus + 3 * mpmath.cos(2 * Kepler_i))
        return (self.coskvim)

    def deorbit_time(self, weight_system, length, coskvim, perigee, apogee, heigth_atmosphere = 120000, Bm =4 * pow(10, -5), R =125 * pow(10, -8)):
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
        self.deltat = ((weight_system * R) / (2 * length * pow(self.Kepler_a, 6) * 1 * coskvim * Bm * Bm)) * (pow(perigee, 7) - pow(heigth_atmosphere, 7))
        self.deltatmes = self.deltat * 3.8052 * pow(10, -7)
        print(f"Время увода системы массой {weight_system} с высоты орбиты {perigee} тросовой системой длинной {length} составляет {self.deltatmes} мес.")

def main():
    list_of_weights = [900, 1800]
    list_of_perigee = [700000, 800000, 900000]
    list_of_apogee = [800000 , 900000, 1000000]
    length_of_tether = [5000, 10000, 15000, 20000, 25000]
    zipped_heights = list(zip(list_of_apogee, list_of_perigee))
    #print(zipped_heights)
    for length in length_of_tether:
        for mass in list_of_weights:
                for tuple_of_heights in zipped_heights:
                    tros = Tros()
                    weigth = tros.weight(length, 0.007, mass)
                    coordinates = tros.coordinates(tuple_of_heights[0])
                    induction = tros.magnetic_induction(coordinates[0],coordinates[1], 60)
                    tros.deorbit_time(weigth, length, induction, tuple_of_heights[0], tuple_of_heights[1])

if __name__ == '__main__':
    main()
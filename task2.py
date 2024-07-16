import math
from typing import Dict, List, Tuple

""" Реализация класса согласно Задания:
    Создайте два класса, Point и Point_creator
    Класс Poit должен иметь поля
        Название
        Координаты словарь с ключами X, Y, Z
    Также у класса должна быть переменная point_count уровня класса,
    которая хранит количество созданных объектов;
    Класс Point_creator должен уметь создавать объект класса Point;
    Также должен иметь методы позволяющие складывать значение координат точек
    и получать новый объект с данными координатами;
    Отнимать координаты точек и получать новый объект с координатами;
    Метод, принимающий на вход базовую точку - объект,
    и список из других точек, возвращать список с расстояниями
    от базовый точки до точек из списка; метод
    (объект_точка,(32,78,156), (132,378,56)), возвращает (32,45)- условно.
"""


class Point:
    point_count = 0

    def __init__(self, name: str, coordinates: Dict[str, float]) -> None:
        self.name = name
        self.coordinates = coordinates
        Point.point_count += 1

    def __str__(self) -> str:
        return f"Point({self.name}, {self.coordinates})"

    def distance_to(self, other: 'Point') -> float:
        # Вычисление расстояния
        dx = self.coordinates['X'] - other.coordinates['X']
        dy = self.coordinates['Y'] - other.coordinates['Y']
        dz = self.coordinates['Z'] - other.coordinates['Z']
        return math.sqrt(dx**2 + dy**2 + dz**2)

    @classmethod
    def get_point_count(cls) -> int:
        return cls.point_count


class PointCreator:
    @staticmethod
    def create_point(name: str,
                     coordinates: Tuple[float, float, float]
                     ) -> Point:
        return Point(name, {'X': coordinates[0],
                            'Y': coordinates[1],
                            'Z': coordinates[2]})

    @staticmethod
    def add_points(point1: Point, point2: Point) -> Point:
        new_coordinates = {
            'X': point1.coordinates['X'] + point2.coordinates['X'],
            'Y': point1.coordinates['Y'] + point2.coordinates['Y'],
            'Z': point1.coordinates['Z'] + point2.coordinates['Z']
        }
        return Point('SumPoint', new_coordinates)

    @staticmethod
    def subtract_points(point1: Point, point2: Point) -> Point:
        new_coordinates = {
            'X': point1.coordinates['X'] - point2.coordinates['X'],
            'Y': point1.coordinates['Y'] - point2.coordinates['Y'],
            'Z': point1.coordinates['Z'] - point2.coordinates['Z']
        }
        return Point('DifferencePoint', new_coordinates)

    @staticmethod
    def distances_from_base(base_point: Point, points: List[Point]
                            ) -> List[float]:
        """ Дистанция вычисляется с использованием евклидового расстояния"""
        distances = []
        for point in points:
            distance = base_point.distance_to(point)
            distances.append(distance)
        return distances


# Примеры использования
if __name__ == "__main__":
    point1 = PointCreator.create_point("Point1", (10, 10, 13))
    point2 = PointCreator.create_point("Point2", (12, 11, 12))
    point3 = PointCreator.create_point("Point3", (13, 14, 17))

    sum_point = PointCreator.add_points(point1, point2)
    diff_point = PointCreator.subtract_points(point3, point1)

    distances = PointCreator.distances_from_base(point1, [point2, point3])

    print(f"Точка1: {point1}")
    print(f"Точка2: {point2}")
    print(f"Точка3: {point3}")
    print(f"Сумма точек: {sum_point}")
    print(f"Разница между точками: {diff_point}")
    print(f"Дистанция от точки1: {distances}")
    print(f"Total Points Created: {Point.get_point_count()}")

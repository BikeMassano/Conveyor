from Model.Constants import MIN_TRANSFER_TIME, MAX_TRANSFER_TIME
from Model.Enums.OperationType import OperationType
from Model.Enums.ProductComplexity import ProductComplexity

class ConveyorLineSegment:
    def __init__(self, operation_type, product_complexity, preform_operation_time, form_operation_time, postform_operation_time, cycle_time, transfer_time):

        if operation_type not in OperationType:
            raise ValueError("Операция не определена!")
        if product_complexity not in ProductComplexity:
            raise ValueError("Сложность изделия не определена!")
        if preform_operation_time < 0:
            raise ValueError("Средняя продолжительность операции на доформочном участке не может быть отрицательной!")
        if form_operation_time < 0:
            raise ValueError("Средняя продолжительность операции на формочном участке не может быть отрицательной!")
        if postform_operation_time < 0:
            raise ValueError("Средняя продолжительность операции на постформочном участке не может быть отрицательной!")
        if cycle_time < 0:
            raise ValueError("Продолжительность цикла формования не может быть отрицательной!")
        if cycle_time <= transfer_time:
            raise ValueError("Продолжительность цикла формования должна быть больше продолжительности передвижения тележек!")
        if MIN_TRANSFER_TIME > transfer_time or transfer_time > MAX_TRANSFER_TIME:
            raise ValueError(f'Продолжительность передвижения тележек должна быть в диапазоне от {MIN_TRANSFER_TIME} до {MAX_TRANSFER_TIME}!')

        self.operation_type = operation_type
        self.product_complexity = product_complexity
        self.preform_operation_time = preform_operation_time
        self.form_operation_time = form_operation_time
        self.postform_operation_time = postform_operation_time
        self.cycle_time = cycle_time
        self.transfer_time = transfer_time
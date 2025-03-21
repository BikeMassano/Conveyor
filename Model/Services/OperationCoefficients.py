"""Сервис для полцучения коэффициентов конвейерной линии"""
class OperationCoefficients:
    def __init__(self, operation_type, product_complexity, coefficient_numerator, coefficient_denominator=None):
        
        self.operation_type = operation_type
        self.product_complexity = product_complexity
        self.coefficient_numerator = coefficient_numerator
        self.coefficient_denominator = coefficient_denominator

    def get_coefficient(self, use_denominator=False):
        
        return self.coefficient_denominator if use_denominator and self.coefficient_denominator is not None else self.coefficient_numerator
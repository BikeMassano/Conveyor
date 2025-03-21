""" Сервис для расчета постов конвейерной линии """
class CalculationService:
    def calculate_posts(self, segment, coefficient, use_denominator=False):
        kh = coefficient.get_coefficient(use_denominator)
        result = ((segment.preform_operation_time + segment.form_operation_time + segment.postform_operation_time) * kh) / (segment.cycle_time - segment.transfer_time)
        if result >= 0:
            return result
        else:
            raise ValueError("Результат расчета не может быть отрицательным!")
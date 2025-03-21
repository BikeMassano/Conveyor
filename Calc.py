from Views.MainView import MainView
from Controllers.CalculatorController import CalculatorController
from Model.Services.CalculationService import CalculationService
from Model.Services.ExportService import ExportService

# Запуск приложения
if __name__ == "__main__":
    calculation_service = CalculationService()
    export_service = ExportService()
    view = MainView(None)
    controller = CalculatorController(view, calculation_service, export_service)
    view._controller = controller
    view.run()
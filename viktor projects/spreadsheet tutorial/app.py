from pathlib import Path
import pandas as pd
import plotly.express as px

from viktor import ViktorController, File
from viktor.parametrization import ViktorParametrization, DownloadButton, LineBreak, NumberField, Section, Tab, Image
from viktor.external.spreadsheet import SpreadsheetCalculation, SpreadsheetCalculationInput
from viktor.result import DownloadResult
from viktor.views import DataGroup, DataItem, DataResult, DataView, PlotlyResult, PlotlyView

class Parametrization(ViktorParametrization):
    general = Tab('General')
    general.beam = Section('Beam')
    general.beam.schematic = Image(path='beam_calcs.png')
    general.beam.length = NumberField('Length (L)', suffix='mm', default=100, max=100)
    general.beam.width = NumberField('Width (W)', suffix='mm', default=10)
    general.beam.height = NumberField('Height (H)', suffix='mm', default=10)
    general.beam.E = NumberField('Modulus of Elasticity (E)', default=200000, suffix='MPa')

    general.loads = Section('Loads')
    general.loads.aw = NumberField('Starting point of load (aw)', suffix='mm', default = 9)
    general.loads.nl = LineBreak()
    general.loads.wa = NumberField('Distributed load amplitude (wa)', suffix = 'N/mm', flex=40, default=5)
    general.loads.wL = NumberField('Distributed load amplitude (wL)', suffix = 'N/mm', flex=40, default=5)

    downloads = Tab('Download')
    downloads.calculation_sheet = Section('Calculation sheet')
    downloads.calculation_sheet.btn = DownloadButton('Download', 'download_spreadsheet')

class Controller(ViktorController):
    label = 'My Entity Type'
    parametrization = Parametrization

    def get_evaled_spreadsheet(self, params):
        inputs = [
            SpreadsheetCalculationInput('L', params.general.beam.length),
            SpreadsheetCalculationInput('W', params.general.beam.width),
            SpreadsheetCalculationInput('H', params.general.beam.height),
            SpreadsheetCalculationInput('E', params.general.beam.E),
            SpreadsheetCalculationInput('aw', params.general.loads.aw),
            SpreadsheetCalculationInput('wa', params.general.loads.wa),
            SpreadsheetCalculationInput('wL', params.general.loads.wL),
        ]
        sheet_path = Path(__file__).parent / 'beam_calculation.xls'
        sheet = SpreadsheetCalculation.from_path(sheet_path, inputs=inputs)
        result = sheet.evaluate(include_filled_file=True)

        return result

    @DataView('Results', duration_guess=1)
    def get_data_view(self, params, **kwargs):
        result = self.get_evaled_spreadsheet(params)

        max_deflection = result.values['maximum_deflection']
        max_bending_stress = result.values['maximum_bending_stress']

        data = DataGroup(
            maximum_deflection=DataItem('Maximum deflection', max_deflection, suffix='micron', number_of_decimals=2),
            maximum_bending_stress=DataItem('Maximum bending stress', max_bending_stress, suffix='MPa', number_of_decimals=2)
        )

        return DataResult(data)

    @PlotlyView('Beam Curvature', duration_guess=1)
    def beam_curvature(self, params, **kwargs):
        result = self.get_evaled_spreadsheet(params)
        evaluated_file = File.from_data(result.file_content)
        with evaluated_file.open_binary() as fp:
            data_df = pd.read_excel(fp, sheet_name='Data')
        deflection_data = data_df['Deflection (microns)'].head(params.general.beam.length+1)
        fig = px.line(deflection_data, title='Beam deflection', labels={'value': 'Deflection (microns)', 'index': 'Length (mm)'})
        return PlotlyResult(fig.to_json())

    def download_spreadsheet(self, params, **kwargs):
        result = self.get_evaled_spreadsheet(params)
        return DownloadResult(result.file_content, 'evaluated_beam.xlsx')
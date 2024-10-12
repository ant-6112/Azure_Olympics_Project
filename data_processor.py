import pandas as pd

class DataProcessor:
    def filter_data(self, data, region, risk_type):
        filtered = data[data['Region'] == region]
        if risk_type != 'All':
            filtered = filtered[filtered['Tax or Financial Risk'] == risk_type]

        filtered_War = filtered[filtered['Due < 42 Days'] == 1]
        filtered_Due = filtered[filtered['Due > 42 Days'] == 1]
        filtered = pd.concat([filtered_War, filtered_Due])
        return filtered

    def get_risk_info(self, risk):
        return f"""
        Steward Name: {risk['Steward Name']}
        Description: {risk['Description']}
        Region: {risk['Region']}
        Risk Type: {risk['Tax or Financial Risk']}
        Due < 42 Days: {'Yes' if risk['Due < 42 Days'] == 1 else 'No'}
        Due > 42 Days: {'Yes' if risk['Due > 42 Days'] == 1 else 'No'}
        """

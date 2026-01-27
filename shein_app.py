import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO


@st.cache_data
def subm_upload(subm_data):
    if subm_data is not None:
        region = pd.read_excel('region.xlsx')
        subm_df = pd.read_csv(subm_data,
                                usecols=['Waybill No.', 'Order Type', 'Client Weight','Delivery Station',
                                         'PPD/COD','COD','Client Volume(cm³)','Create Operator',],
                                         dtype= str)
        subm_df = subm_df.merge(region, on= 'Delivery Station', how='left')
        
        return subm_df
    return None

def add_cols(subm_df):
    if subm_df is not None:
        # filter shein only (excluding Reverse)
        filt_df = subm_df[(subm_df['Create Operator'].str.contains('shein',case=False,na=False)) &
                   (subm_df['Order Type']!='Reverse Pickup(Return & Refund)')]
        # Route column
        filt_df['Order Route'] = np.where(filt_df['Create Operator'].str.contains('road', case=False, na=False),
                                                                                    'By Road', 'By Air')
        # Size column
        filt_df['Order Size'] = np.where(filt_df['Client Volume(cm³)']==0.01,'Pouch','Box')

        # Order Value
        filt_df['Order Value'] = np.where(filt_df['COD']<999,'LV', 'HV')

        return filt_df
    
def dfs_creation(filt_df):
    if filt_df is None or filt_df.empty:
        return (pd.DataFrame(),) * 6
    require_arrang = ['West Region', 'Central Region','South Region','North Region','East Region']
    # Regional Total
    reg_total = filt_df.groupby('Region').size().to_frame(name='Count').T
    reg_total = reg_total[require_arrang]
    reg_total['Total'] = reg_total.sum(axis=1)

    # Regional Weight
    reg_weight = filt_df.groupby('Region')['Client Weight'].mean().to_frame().T
    reg_weight = reg_weight[require_arrang]
    reg_weight['Total'] = filt_df['Client Weight'].mean()
    reg_weight = reg_weight.round(2)

    #Regional Route
    reg_route = pd.crosstab(filt_df['Order Route'], filt_df['Region'])
    reg_route = reg_route[require_arrang]
    reg_route['Total'] = reg_route.sum(axis=1)

    # Region Size
    reg_size = pd.crosstab(filt_df['Order Size'], filt_df['Region'], normalize='columns')
    reg_size = reg_size[require_arrang]
    reg_size['Total'] =filt_df['Order Size'].value_counts(normalize=True)
    reg_size = reg_size.round(4)

    #Region order value
    reg_value = pd.crosstab(filt_df['Order Value'], filt_df['Region'], normalize='columns')
    reg_value = reg_value[require_arrang]
    reg_value['Total'] = filt_df['Order Value'].value_counts(normalize=True)
    reg_value = reg_value.round(4)

    # Delivery Method
    reg_d_method = pd.crosstab(filt_df['PPD/COD'], filt_df['Region'], normalize='columns')
    reg_d_method = reg_d_method[require_arrang]
    reg_d_method['Total'] = filt_df['PPD/COD'].value_counts(normalize=True)
    reg_d_method = reg_d_method.round(4)

    return reg_total,reg_weight,reg_route,reg_size,reg_value,reg_d_method

def data_download(reg_total,reg_weight,reg_route,reg_size,reg_value,reg_d_method):
    if reg_total is not None:
        dfs = [reg_total,reg_weight,reg_route,reg_size,reg_value,reg_d_method]
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            startrow = 1
            for df in dfs:
                df.to_excel(writer, sheet_name="Summary", startrow=startrow)
                startrow += len(df) + 3
        output.seek(0)
        return output


def main():
    st.set_page_config(page_title="Shein App",
                        page_icon=':alarm_clock:',
                        layout="wide",
                        initial_sidebar_state="expanded")
    
    with st.container(): 
        st.markdown("### Shein Data Automation")
        col1, col2, col3 = st.columns([2,0.25,2])
        with col1:
            subm_data = st.file_uploader("Upload Submission File(CSV format(comma delimited))", type="csv", key="submission")
            subm_df = subm_upload(subm_data)
            filt_df = add_cols(subm_df)
            reg_total,reg_weight,reg_route,reg_size,reg_value,reg_d_method = dfs_creation(filt_df)
        with col3:
            st.markdown('#### Uploaded file info')
            if subm_data is not None:
                st.write(f"Total submission is: {subm_df['Waybill No.'].nunique()}")
                st.write(f"Shein count excluding Reverse: {filt_df['Waybill No.'].nunique()}")
    st.markdown('---')

    with st.container():
        # st.markdown("### Download Data")
        output = data_download(reg_total,reg_weight,reg_route,reg_size,reg_value,reg_d_method)
        st.download_button(
                            label="Download Output file",
                            data=output,
                            file_name="shein_output.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
    st.markdown('---')

    with st.container():
        col1,col2,col3 = st.columns([0.25,2,0.25])
        if subm_data is not None:
            with col2:
                st.markdown("#### Regional Total")
                st.dataframe(reg_total)
                st.markdown('---')

                st.markdown("#### Region-wise Weight")
                st.dataframe(reg_weight)
                st.markdown('---')

                st.markdown("#### Shipment Route")
                st.dataframe(reg_route)
                st.markdown('---')

                st.markdown("#### Region-wise Size")
                st.dataframe(reg_size)
                st.markdown('---')

                st.markdown("#### Shipment value")
                st.dataframe(reg_value)
                st.markdown('---')
                
                st.markdown("#### Delivery Method")
                st.dataframe(reg_d_method)
                st.markdown('---')
                
    # st.dataframe()

if __name__ == "__main__":

    main()

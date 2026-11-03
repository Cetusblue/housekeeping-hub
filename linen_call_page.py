"""Linen Call Log UI and monthly Excel export."""
from io import BytesIO
from datetime import datetime
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from master_loader import load_linen_locations_rows, load_linen_master_rows, get_linen_bundle_map
from linen_call_db import create_call, list_calls, get_call, delete_call

# Operational shortlist: order is intentional.
PRIORITY_SETS = ('PYJ_F_5XL', 'PYJ_F_10XL', 'PYJ_M_5XL', 'PYJ_M_10XL')
PRIORITY_ITEMS = ('BLANKET FLANNEL FS', 'DRAW SHEET FS', 'BATH TOWEL FS', 'FITTED BED SHEET FS')
OTHER_SETS = ('PYJ_F_L', 'PYJ_F_XL', 'PYJ_M_L', 'PYJ_M_XL')
OTHER_ITEMS = ('ORTHO PANT FS', 'URO PANT FS', 'PATIENT GOWN FS', 'SARONG FS',
               'FEMALE KIMONO FS', 'MALE KIMONO FS', 'KIMONO FS',
               'X RAY GOWN  M', 'X RAY GOWN  L', 'X RAY GOWN  3XL',
               'TABLE SKIRTING FS')

def _home():
    st.session_state['page']='home'
    st.rerun()

def _departments():
    values=set()
    for r in load_linen_locations_rows():
        t,l,z=(str(r.get(k) or '').strip() for k in ('tower','level','zone'))
        if t and l: values.add((t,l,z))
    return sorted(values)

def _items():
    bundles = get_linen_bundle_map()
    individual = {r['item_name'].upper().strip(): r for r in load_linen_master_rows()}
    result = []

    def add_set(bundle_id):
        if bundle_id not in bundles:
            raise ValueError(f'Linen Call Log bundle not found: {bundle_id}')
        result.append({'item_key': 'B:' + bundle_id, 'item_name': bundles[bundle_id]['bundle_name']})

    def add_item(name):
        if name not in individual:
            raise ValueError(f'Linen Call Log item not found: {name}')
        row = individual[name]
        result.append({'item_key': 'I:' + str(row['item_no']), 'item_name': row['item_name']})

    for bid in PRIORITY_SETS: add_set(bid)
    for name in PRIORITY_ITEMS: add_item(name)
    priority_count = len(result)
    for bid in OTHER_SETS: add_set(bid)
    for name in OTHER_ITEMS: add_item(name)
    return result, priority_count

def _excel(month):
    wb=Workbook();ws=wb.active;ws.title='Linen Call Log'
    headers=['Request ID','Timestamp (SGT)','Tower','Level','Zone','Item','Request Qty','Delivered Qty','Recorded By']
    ws.append(headers)
    for call in reversed(list_calls(month=month)):
        _,lines=get_call(call['id'])
        for line in lines:
            ws.append([call['id'],call['created_at'],call['tower'],call['level'],call['zone'],line['item_name'],line['request_qty'],line['delivered_qty'],call['created_by']])
    for cell in ws[1]:
        cell.font=Font(color='FFFFFF',bold=True)
        cell.fill=PatternFill('solid',fgColor='214A68')
        cell.alignment=Alignment(wrap_text=True)
    ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
    for col,width in {'A':15,'B':24,'C':12,'D':12,'E':18,'F':38,'G':17,'H':18,'I':22}.items():ws.column_dimensions[col].width=width
    data=BytesIO();wb.save(data);return data.getvalue()

def page_linen_call_log():
    user=st.session_state.get('user') or {}
    if user.get('role') not in ('LINTEAM','LINSUP'):
        st.error('Access denied.');return
    st.title('Linen Call Log')
    st.caption('Record requested and delivered quantities together. This does not deduct inventory.')
    if st.button('← Back to Home'):_home()
    departments=_departments()
    if not departments:
        st.warning('No departments found in Linen Location Master.');return
    labels=[f'{t} / {l} / {z or "—"}' for t,l,z in departments]
    with st.form('linen_call_entry',clear_on_submit=True):
        department=st.selectbox('Department (Tower / Level / Zone)',range(len(departments)),format_func=lambda i:labels[i])
        st.subheader('Priority Items')
        st.caption('Enter whole numbers. Zero delivered is allowed.')
        lines=[]
        items, priority_count = _items()
        for index,item in enumerate(items):
            if index == priority_count:
                st.subheader('Other Linen Items')
            a,b,c=st.columns([4,1,1])
            with a:st.write(item['item_name'])
            with b:req=st.number_input('Request',min_value=0,step=1,key='lcr_'+item['item_key'],label_visibility='visible')
            with c:deliv=st.number_input('Delivered',min_value=0,step=1,key='lcd_'+item['item_key'],label_visibility='visible')
            lines.append(dict(item,request_qty=int(req),delivered_qty=int(deliv)))
        submitted=st.form_submit_button('Submit Linen Call',type='primary',use_container_width=True)
    if submitted:
        try:
            call_id=create_call(*departments[department],user['username'],lines)
            st.success(f'Linen Call #{call_id} saved.')
        except ValueError as e:st.warning(str(e))
        except Exception as e:st.error(f'Could not save: {e}')
    st.divider();st.subheader('Latest 30 Active Requests')
    recent=list_calls(limit=30)
    if recent:
        st.dataframe([{'Request ID':r['id'],'Timestamp (SGT)':r['created_at'],'Tower':r['tower'],'Level':r['level'],'Zone':r['zone'],'Recorded By':r['created_by']} for r in recent],hide_index=True,use_container_width=True)
    else:st.info('No active records yet.')
    st.subheader('Delete a Record')
    rid=st.number_input('Request ID',min_value=1,step=1,key='lc_delete_id')
    if st.button('Preview Request',key='lc_preview'):
        st.session_state['lc_preview_id']=int(rid)
    if st.session_state.get('lc_preview_id'):
        record,details=get_call(st.session_state['lc_preview_id'])
        if not record or record['deleted_at']:
            st.warning('Request not found or already deleted.')
        else:
            st.write(f"Request #{record['id']} — {record['tower']} / {record['level']} / {record['zone']} — {record['created_at']} SGT")
            st.dataframe([{'Item':r['item_name'],'Request':r['request_qty'],'Delivered':r['delivered_qty']} for r in details],hide_index=True,use_container_width=True)
            confirm=st.checkbox('I confirm this entire record should be deleted.',key='lc_confirm')
            if st.button('Delete Request',disabled=not confirm,type='secondary'):
                if delete_call(record['id'],user['username']):
                    st.session_state.pop('lc_preview_id',None)
                    st.success('Record deleted. It will be excluded from monthly reports.')
                    st.rerun()
                else:st.warning('Record was already deleted.')
    st.divider();st.subheader('Monthly Report')
    month=st.text_input('Month (YYYY-MM)',value=datetime.now().strftime('%Y-%m'))
    try:
        datetime.strptime(month,'%Y-%m')
        st.download_button('Download Excel Report',data=_excel(month),file_name=f'Linen_Call_Log_{month}.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except ValueError:st.warning('Enter a valid month in YYYY-MM format.')

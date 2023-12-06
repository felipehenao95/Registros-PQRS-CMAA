import os
import dash
from dash import dcc, html, Input, Output, State
from docxtpl import DocxTemplate
from dash import dash_table
import pandas as pd
from datetime import datetime as dt
from datetime import date
from io import BytesIO
import base64



meses_espanol = {1: 'ENERO',2: 'FEBRERO',3: 'MARZO',4: 'ABRIL',5: 'MAYO',
                        6: 'JUNIO',7: 'JULIO',8: 'AGOSTO',9: 'SEPTIEMBRE',10: 'OCTUBRE',
                        11: 'NOVIEMBRE',12: 'DICIEMBRE'}

fecha_actual = dt.now().date()

output_dataframe = pd.DataFrame(columns=['fecha_llegada','radicado', 'fecha_radicado','nombre',
                                        'correo','direccion','barrio','localidad','tipo', 'asunto'])
## COLORES:
# Fondo app: #c8dbec
# Fondo fechas: 
# Fondo Inputs: #ecf2f893


# Crear la aplicación Dash
app = dash.Dash(__name__)
server = app.server
# Diseño del formulario
app.layout = html.Div(style={'backgroundColor': '#c8dbec'}, children=[
    html.H1("FORMULARIO DE REGISTRO DE PQRS Y GENERADOR DE PROYECCIÓN DE RTA", 
            className="app-header"),
    html.Div([
        html.Div([
            html.Div(style={'display': 'flex', 'flexDirection': 'row', 'padding': 10}, className="section-label",children=[
                html.Label('Fecha de Llegada Applus+ K2:', style={'fontFamily': 'Calibri', 'margin-right': '77px'}),
                dcc.DatePickerSingle(id='input-fecha-llegada',min_date_allowed=date(2022, 1, 1),
                                    max_date_allowed=date(2025, 12, 31),
                                    initial_visible_month=fecha_actual,date=fecha_actual,
                                    style={'backgroundColor': '#ecf2f893'}),
            ]),
            html.Div(style={'display': 'flex', 'flexDirection': 'row', 'padding': 10}, className="section-label", children=[
                html.Label('Fecha del oficio (RTA):', style={'fontFamily': 'Calibri','margin-right': '119px'}),
                dcc.DatePickerSingle(id='input-fecha-oficio2',min_date_allowed=date(2022, 1, 1),
                                    max_date_allowed=date(2025, 12, 31),
                                    initial_visible_month=fecha_actual,date=fecha_actual,
                                    style={'backgroundColor': '#ecf2f893'}),
            ]),
            html.Div(style={'display': 'flex', 'flexDirection': 'row', 'padding': 10}, className="section-label", children=[
                html.Label('Fecha del Radicado:', style={'fontFamily': 'Calibri','margin-right': '130px'}),
                dcc.DatePickerSingle(id='input-fecha-radicado2',min_date_allowed=date(2022, 1, 1),
                                    max_date_allowed=date(2025, 12, 31),
                                    initial_visible_month=fecha_actual,date=fecha_actual,
                                    style={'backgroundColor': 'lightgrey'}),
            ]),
            dcc.Dropdown(id='dropdown-saludo',options=[
                {'label': 'Señor', 'value': 'Señor'},
                {'label': 'Señora', 'value': 'Señora'}
            ],value='Señor',
            style={'backgroundColor': '#ecf2f893'},
            className="section-dropdown"),

            html.Br(),
            dcc.Input(id='input-nombre', type='text', placeholder='Nombre Completo', style={'width': '400px'}, className='section-text-input'),
            html.Br(),
            dcc.Input(id='input-radicado', type='text', placeholder='Numero de radicado', style={'width': '400px'}, className='section-text-input'),
            html.Br(),
            dcc.Input(id='input-correo', type='text', placeholder='Correo', style={'width': '400px'}, className='section-text-input'),
            html.Br(),
            dcc.Input(id='input-direccion', type='text', placeholder='Direccion', style={'width': '400px'}, className='section-text-input'),
            html.Br(),
            dcc.Input(id='input-latitud', type='text', placeholder='Latitud', style={'width': '400px'}, className='section-text-input'),
            html.Br(),
            dcc.Input(id='input-longitud', type='text', placeholder='Longitud', style={'width': '400px'}, className='section-text-input'),
            html.Br(),
            dcc.Input(id='input-barrio', type='text', placeholder='Barrio', style={'width': '400px'}, className='section-text-input'),
            html.Br(),
            dcc.Input(id='input-localidad', type='text', placeholder='Localidad', style={'width': '400px'}, className='section-text-input'),
            html.Br(),
            dcc.Textarea(id='input-asunto', placeholder='Asunto', 
                        style={'width': '400px', 'height': '140px', 'whiteSpace': 'pre-wrap'}, className='section-text-tarea'),
            html.Br(),
            html.A(
                id='download-link',
                children='No hay archivos para descargar',
                href='',
                target='_blank',
                download='',
                style={'text-align': 'center'}
            ),
        ], style={'display': 'flex', 'flexDirection': 'column', 'margin-right': '30px'}),
        html.Div([
            html.Label('Seleccione el tipo de DP:', style={'fontFamily': 'Calibri', 'fontWeight': 'bold'}),

            dcc.Dropdown(id='dropdown-tipoDP',options=[
                {'label': 'Queja', 'value': 'Queja'},
                {'label': 'Solicitud', 'value': 'Solicitud'}
            ],value='Queja',
            className="section-dropdown"),
            html.Br(),
            html.Label('Seleccione el tema del DP:', style={'fontFamily': 'Calibri', 'fontWeight': 'bold'}),

            dcc.Dropdown(id='dropdown-temaDP',options=[
                {'label': 'Ruido', 'value': 'Ruido'},
                {'label': 'Rutas Aereas o Altitud', 'value': 'Rutas Aereas o Altitud'},
                {'label': 'Visita o Socialización', 'value': 'Visitas o Socialización'},
                {'label': 'Insonorización', 'value': 'Insonosización'}
            ],value='Ruido',
            className="section-dropdown"),
            html.Br(),
            dcc.Textarea(id='input-peticion', placeholder='Petición', 
                        style={'width': '600px', 'height': '100px', 'whiteSpace': 'pre-wrap'},
                        className='section-text-tarea'),
            html.Br(),
            
            dcc.Textarea(id='input-peticion-puntual1', placeholder='Petición Puntual 1', 
                        style={'width': '600px', 'height': '50px', 'whiteSpace': 'pre-wrap'},
                        className='section-text-tarea'),
            html.Br(),
            dcc.Textarea(id='input-peticion-puntual2', placeholder='Petición Puntual 2', 
                        style={'width': '600px', 'height': '50px', 'whiteSpace': 'pre-wrap'},
                        className='section-text-tarea'),
            html.Br(),
            dcc.Textarea(id='input-peticion-puntual3', placeholder='Petición Puntual 3', 
                        style={'width': '600px', 'height': '50px', 'whiteSpace': 'pre-wrap'},
                        className='section-text-tarea'),
            html.Br(),
            dcc.Textarea(id='input-peticion-puntual4', placeholder='Petición Puntual 4', 
                        style={'width': '600px', 'height': '50px', 'whiteSpace': 'pre-wrap'},
                        className='section-text-tarea'),
            html.Br(),
            html.Button( id='exportar-button', n_clicks=0),
            html.Br(),
            html.Button( id='subir-button', n_clicks=0),
            html.Br(),
            html.Button( id='exportar-registros-button', n_clicks=0),
            dcc.Download(id="download-registros"),
            html.Br(),
            html.Button( id='limpiar-button', n_clicks=0),
            html.Br(),
            html.Button('Borrar Registros', id='borrar-registros-button', n_clicks=0)
            
        ], style={'display': 'flex', 'flexDirection': 'column'}),
    ], style={'display': 'flex', 'flexDirection': 'row', 'padding': 10}),
    
    dash_table.DataTable(
        id='tabla-datos',
        columns=[
            {'name': 'Fecha Llegada', 'id': 'fecha_llegada'},
            {'name': 'Radicado', 'id': 'radicado'},
            {'name': 'Fecha Radicado', 'id': 'fecha_radicado'},
            {'name': 'Nombre', 'id': 'nombre'},
            {'name': 'Correo', 'id': 'correo'},
            {'name': 'Dirección', 'id': 'direccion'},
            {'name': 'Barrio', 'id': 'barrio'},
            {'name': 'Localidad', 'id': 'localidad'},
            {'name': 'Tipo', 'id': 'tipo'},
            {'name': 'Asunto', 'id': 'asunto'},
        ],
        style_table={'height': '300px', 'overflowY': 'auto'},
        style_header={
            'textAlign': 'center',
            'backgroundColor': 'lightgrey',  # Puedes cambiar el color de fondo según tus preferencias
            'fontWeight': 'bold',  # Puedes aplicar negrita u otros estilos de fuente
            'font_size': '14px'  # Puedes ajustar el tamaño de la fuente
        },
        style_cell={
            'minWidth': '50px', 'maxWidth': '140px',
            'whiteSpace': 'nowrap',
            'overflow': 'hidden',
            'textOverflow': 'ellipsis',
        },
    ),
])

# Callback para generar el documento al hacer clic en el botón
@app.callback(
    Output('download-link', 'href'),
    Output('download-link', 'download'),
    Output('download-link', 'children'),
    Output("download-registros", "data"),
    Output("exportar-button", "children"),
    Output("subir-button", "children"),
    Output("exportar-registros-button", "children"),
    Output('tabla-datos', 'data'),
    [Input("exportar-button", "n_clicks")],
    [State('dropdown-saludo', 'value'),
    State('input-nombre', 'value'),
    State('input-fecha-llegada', 'date'),
    State('input-fecha-oficio2', 'date'),
    State('input-radicado', 'value'),
    State('input-fecha-radicado2', 'date'),
    State('input-correo', 'value'),
    State('input-direccion', 'value'),
    State('input-latitud', 'value'),
    State('input-longitud', 'value'),
    State('input-barrio', 'value'),
    State('input-localidad', 'value'),
    State('input-asunto', 'value'),
    State('dropdown-tipoDP','value'),
    State('dropdown-temaDP','value'),
    State('input-peticion', 'value'),
    State('input-peticion-puntual1', 'value'),
    State('input-peticion-puntual2', 'value'),
    State('input-peticion-puntual3', 'value'),
    State('input-peticion-puntual4', 'value')],
    Input("subir-button", "n_clicks"),
    Input("exportar-registros-button", "n_clicks"),
    Input("borrar-registros-button", "n_clicks")
)
def generar_documento_callback(n_clicks, genero, nombre, fecha_llegada,fecha_oficio2, radicado, 
                                fecha_radicado2, correo, direccion, latitud, longitud, barrio, localidad, asunto, 
                                tipoDP, temaDP, peticion, peticion_puntual1, peticion_puntual2,
                                peticion_puntual3,peticion_puntual4, n_clicks2, n_clicks3, n_clicks4):
    global output_dataframe
    ctx = dash.callback_context
    triggered_id = ctx.triggered_id if ctx.triggered_id else 'nada'

    if 'exportar-button' in triggered_id and nombre:

        fecha_rad = dt.strptime(fecha_radicado2, '%Y-%m-%d')
        dia = fecha_rad.day
        mes = fecha_rad.month
        anio = fecha_rad.year
        mes_espanol = meses_espanol.get(mes).upper()
        fechaRadicado = f'{dia} DE {mes_espanol} DEL {anio}'

        fecha_of = dt.strptime(fecha_oficio2, '%Y-%m-%d')
        dia = fecha_of.day
        mes = fecha_of.month
        anio = fecha_of.year
        mes_espanol = meses_espanol.get(mes).lower()
        fechaOficio = f'{dia} de {mes_espanol} del {anio}'
        
        datos_usuario = {
            'genero_titulo': genero,
            'genero': genero.lower(),
            'nombre': nombre,
            'nombre_titulo': nombre.upper(),
            'fecha_oficio': fechaOficio,
            'radicado': radicado,
            'correo': correo,
            'fecha_radicado': fechaRadicado,
            'direccion': direccion,
            'barrio': barrio,
            'localidad': localidad,
            'asunto': asunto,
            'peticion': peticion,
            'peticion_puntual1': peticion_puntual1,
            'peticion_puntual2': peticion_puntual2,
            'peticion_puntual3': peticion_puntual3,
            'peticion_puntual4': peticion_puntual4,
        }

        plantilla_path = "PlantillaPQRS.docx"
        doc = DocxTemplate(plantilla_path)

        doc.render(datos_usuario)

        output = BytesIO()
        nombre_archivo = f'RTA {nombre}.docx' 
        doc.save(output)
        output.seek(0)
        
        return (f'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{base64.b64encode(output.read()).decode()}',
                nombre_archivo,
                f'Descargar {nombre_archivo}',
                None,
                'Se exportó correctamente' ,
                'Cargar formulario de registro',
                'Exportar Registros (DatosDPs.xlsx)', 
                output_dataframe.to_dict('records') )

    elif 'subir-button' in triggered_id and nombre:
    
        datos_seguimiento = {
            'fecha_llegada': fecha_llegada,
            'radicado': radicado,
            'fecha_radicado': fecha_radicado2,
            'nombre': nombre,
            'correo': correo,
            'direccion': direccion,
            'barrio': barrio,
            'localidad': localidad,
            'tipo': tipoDP,
            'asunto': temaDP,
            'peticion': peticion,
            'latitud': latitud,
            'longitud': longitud,
        }
        df_seg=pd.DataFrame([datos_seguimiento])
        output_dataframe = pd.concat([output_dataframe, df_seg], axis=0)

        return ('','', f'Descarga no disponible',
                None,
                'Exportar Formato de RTA.docx', 
                'Cargar formulario de registro', 
                'Exportar Registros (DatosDPs.xlsx)', 
                output_dataframe.to_dict('records'))

    elif 'exportar-registros-button' in triggered_id:
        fecha_of = dt.strptime(fecha_oficio2, '%Y-%m-%d')
        dia = fecha_of.day
        mes = fecha_of.month
        anio = fecha_of.year
        mes_espanol = meses_espanol.get(mes).lower()
        fechaOficio = f'{dia} de {mes_espanol} del {anio}'
        # output_dataframe.to_excel('DatosDPs.xlsx', index=False)
        return ('','', f'Descarga no disponible',
                dcc.send_data_frame(output_dataframe.to_excel, f"Registros {fechaOficio}.xlsx"),
                'Exportar Formato de RTA.docx', 
                'Cargar formulario de registro',
                'Registros Exportados (DatosDPs.xlsx)', 
                output_dataframe.to_dict('records'))
    
    elif 'borrar-registros-button' in triggered_id:
        output_dataframe = output_dataframe.drop(index=output_dataframe.index)
        return ('','', f'Descarga no disponible', None,
                'Exportar Formato de RTA.docx', 
                'Cargar formulario de registro',
                'Registros Exportados (DatosDPs.xlsx)', 
                output_dataframe.to_dict('records'))
    
    else:
        
        return ('','', f'Descarga no disponible', None,
                'Exportar Formato de RTA.docx', 
                'Cargar formulario de registro', 
                'Exportar Registros (DatosDPs.xlsx)', 
                output_dataframe.to_dict('records'))


@app.callback(
    Output('limpiar-button', 'children'),
    Output('dropdown-saludo', 'value'),
    Output('input-nombre', 'value'),
    Output('input-fecha-llegada', 'date'),
    Output('input-fecha-oficio2', 'date'),
    Output('input-radicado', 'value'),
    Output('input-fecha-radicado2', 'date'),
    Output('input-correo', 'value'),
    Output('input-direccion', 'value'),
    Output('input-latitud', 'value'),
    Output('input-longitud', 'value'),
    Output('input-barrio', 'value'),
    Output('input-localidad', 'value'),
    Output('input-asunto', 'value'),
    Output('dropdown-tipoDP','value'),
    Output('dropdown-temaDP','value'),
    Output('input-peticion', 'value'),
    Output('input-peticion-puntual1', 'value'),
    Output('input-peticion-puntual2', 'value'),
    Output('input-peticion-puntual3', 'value'),
    Output('input-peticion-puntual4', 'value'),
    Input("limpiar-button", "n_clicks"),
)

def limpiar_campos_callback(n_clicks):

    if n_clicks is None:
        # Sin clic, no hay cambios
        raise dash.exceptions.PreventUpdate
    
    return ('Limpiar Campos','Señor','',fecha_actual,fecha_actual,'',fecha_actual,'','','','','','','','Queja','Ruido','','','','','')


if __name__ == '__main__':
    app.run_server(debug=True)

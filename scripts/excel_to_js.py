import pandas as pd
import json
from datetime import datetime
import os

def excel_to_js(excel_file, output_file="data.js"):
    """
    Convierte un archivo Excel a data.js con configuración completa
    """
    
    print(f"📂 Leyendo archivo: {excel_file}")
    excel_data = pd.read_excel(excel_file, sheet_name=None)
    
    # ============================================
    # 1. LEER CONFIGURACIÓN GENERAL
    # ============================================
    config = {}
    if 'CONFIGURACION' in excel_data:
        df_config = excel_data['CONFIGURACION']
        for _, row in df_config.iterrows():
            if pd.notna(row.get('Parametro')):
                config[str(row['Parametro'])] = str(row['Valor']) if pd.notna(row.get('Valor')) else ''
    
    # Valores por defecto si no existen
    fecha_corte = config.get('Fecha Corte', datetime.now().strftime('%Y-%m-%d'))
    global_inicio = config.get('Proyecto Global Inicio', '2025-05-01')
    global_fin = config.get('Proyecto Global Fin', '2026-04-30')
    
    # ============================================
    # 2. LEER CONFIGURACIÓN DE PROYECTOS
    # ============================================
    proyectos = []
    if 'PROYECTOS' in excel_data:
        df_proyectos = excel_data['PROYECTOS']
        for _, row in df_proyectos.iterrows():
            if pd.notna(row.get('Proyecto')):
                proyecto = {
                    'nombre': str(row['Proyecto']),
                    'owner': str(row['Dueño']),
                    'tipo_calculo': str(row['Tipo Calculo']).upper(),
                    'inicio': str(row['Inicio']),
                    'fin': str(row['Fin']),
                    'color': str(row['Color']) if pd.notna(row.get('Color')) else '#38bdf8'
                }
                proyectos.append(proyecto)
    
    # ============================================
    # 3. LEER DATOS DE ACTIVIDADES
    # ============================================
    todas_actividades = []
    
    # Hojas que contienen actividades (excluir CONFIGURACION y PROYECTOS)
    hojas_actividades = [h for h in excel_data.keys() if h not in ['CONFIGURACION', 'PROYECTOS']]
    
    for sheet_name in hojas_actividades:
        print(f"📊 Procesando hoja: {sheet_name}")
        df = excel_data[sheet_name]
        
        # Buscar el proyecto al que pertenece esta hoja
        proyecto_asociado = next((p for p in proyectos if p['nombre'] == sheet_name), None)
        owner = proyecto_asociado['owner'] if proyecto_asociado else sheet_name
        
        actividades = []
        for _, row in df.iterrows():
            if pd.notna(row.get('ID')) and pd.notna(row.get('Actividad')):
                actividad = {
                    'id': str(row['ID']),
                    'reporte': str(row['Actividad']),
                    'fase': str(row['Fase']) if pd.notna(row.get('Fase')) else sheet_name,
                    'owner': owner,
                    'inicio': str(row['Inicio']),
                    'fin': str(row['Fin']),
                    'estado': str(row['Estado']) if pd.notna(row.get('Estado')) else 'En progreso',
                    'avance': float(row['Avance']) if pd.notna(row.get('Avance')) else 0.0
                }
                
                # Agregar peso si existe (para proyectos ponderados)
                if pd.notna(row.get('Peso')):
                    actividad['peso'] = float(row['Peso'])
                
                actividades.append(actividad)
        
        todas_actividades.extend(actividades)
        print(f"   → {len(actividades)} actividades")
    
    # ============================================
    # 4. CREAR ARCHIVO JS
    # ============================================
    js_content = f"""// ============================================
// ARCHIVO GENERADO AUTOMÁTICAMENTE
// Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
// Fuente: {os.path.basename(excel_file)}
// ============================================

// CONFIGURACIÓN GENERAL
const CONFIG = {{
    FECHA_CORTE: "{fecha_corte}",
    GLOBAL_INICIO: "{global_inicio}",
    GLOBAL_FIN: "{global_fin}"
}};

// CONFIGURACIÓN DE PROYECTOS
const proyectosConfig = {json.dumps(proyectos, indent=2, ensure_ascii=False)};

// DATOS DE ACTIVIDADES
const actividadesData = {json.dumps(todas_actividades, indent=2, ensure_ascii=False)};

// ============================================
// NO MODIFICAR MANUALMENTE
// ============================================
"""
    
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(js_content)
    
    print(f"\n✅ Archivo {output_file} generado correctamente")
    print(f"   Fecha de corte: {fecha_corte}")
    print(f"   Proyectos: {len(proyectos)}")
    print(f"   Actividades: {len(todas_actividades)}")
    
    return config, proyectos, todas_actividades

def create_template_excel(output_file="plantilla_proyecto.xlsx"):
    """
    Crea un archivo Excel de plantilla con todas las hojas necesarias
    """
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Hoja CONFIGURACION
        config_df = pd.DataFrame({
            'Parametro': ['Fecha Corte', 'Proyecto Global Inicio', 'Proyecto Global Fin'],
            'Valor': ['2026-03-12', '2025-05-08', '2026-04-09']
        })
        config_df.to_excel(writer, sheet_name='CONFIGURACION', index=False)
        
        # Hoja PROYECTOS
        proyectos_df = pd.DataFrame({
            'Proyecto': ['TECSUR OBRA', 'SPC GESTIÓN'],
            'Dueño': ['TECSUR', 'SPC'],
            'Tipo Calculo': ['PROMEDIO', 'PONDERADO'],
            'Inicio': ['2026-01-16', '2025-05-08'],
            'Fin': ['2026-04-09', '2026-04-09'],
            'Color': ['#3b82f6', '#a855f7']
        })
        proyectos_df.to_excel(writer, sheet_name='PROYECTOS', index=False)
        
        # Hoja TECSUR
        tecsur_df = pd.DataFrame({
            'ID': ['R1', 'R2.1', 'R2.2', 'R2.3', 'R2.4', 'R2.5', 'R3.1', 'R3.2', 'R3.3', 'R3.4', 'R3.5', 'R3.6', 'R3.7', 'R3.8', 'R4'],
            'Actividad': [
                'R1 - OBRAS PROVISIONALES Y PRELIMINARES',
                'R2.1 - MONTAJE DE POSTES - MEDIA TENSIÓN',
                'R2.2- MONTAJE DE ARMADOS MT',
                'R2.3 - MONTAJE DE CONDUCTORES',
                'R2.4 - INSTALACION DE PUESTA A TIERRA',
                'R2.5 - TRABAJOS SUBTERRÁNEOS MT',
                'R3.1 - TRABAJOS SUBTERRANEOS BT',
                'R3.2 - MONTAJE DE POSTES -BT',
                'R3.3 - MONTAJE DE ARMADOS BT',
                'R3.4 - MONTAJE DE CONDUCTORES BT',
                'R3.5 - INSTALACION DE PUESTA A TIERRA BT',
                'R3.6 - INSTALACION DE EQUIPOS ALUMBRADO',
                'R3.7 - INSTALACION DE EMPALMES Y CONECTORES',
                'R3.8 - CONEXIÓN CAJA DISTRIBUCIÓN',
                'R4 - PRUEBAS ELÉCTRICAS Y ENTREGA'
            ],
            'Fase': ['Fase 1', 'Fase 1', 'Fase 1', 'Fase 1', 'Fase 1', 'Fase 1', 
                    'Fase 2', 'Fase 2', 'Fase 2', 'Fase 2', 'Fase 2', 'Fase 2', 'Fase 2', 'Fase 2', 'Fase 2'],
            'Owner': ['TECSUR'] * 15,
            'Inicio': ['2026-01-31', '2026-01-29', '2026-02-09', '2026-02-18', '2026-02-05', '2026-01-16',
                      '2026-02-19', '2026-02-03', '2026-02-19', '2026-03-05', '2026-03-09', '2026-03-10',
                      '2026-03-13', '2026-03-20', '2026-04-08'],
            'Fin': ['2026-04-06', '2026-03-30', '2026-03-10', '2026-02-26', '2026-03-14', '2026-03-03',
                   '2026-03-19', '2026-03-22', '2026-03-22', '2026-03-24', '2026-03-24', '2026-03-26',
                   '2026-03-26', '2026-03-29', '2026-04-08'],
            'Estado': ['En progreso', 'Terminado', 'En progreso', 'Terminado', 'En progreso', 'Terminado',
                      'En progreso', 'Terminado', 'En progreso', 'En progreso', 'En progreso', 'En progreso',
                      'No iniciado', 'No iniciado', 'No iniciado'],
            'Avance': [0.85, 1.00, 0.87, 1.00, 0.75, 1.00, 0.72, 1.00, 0.96, 0.25, 0.63, 0.64, 0.00, 0.00, 0.00]
        })
        tecsur_df.to_excel(writer, sheet_name='TECSUR', index=False)
        
        # Hoja SPC
        spc_df = pd.DataFrame({
            'ID': ['P1-F1-01', 'P1-F1-02', 'P1-F1-03', 'P1-F1-04', 'P1-F1-05', 'P1-F1-06', 
                   'P1-F1-07', 'P1-F1-08', 'P1-F1-09', 'P1-F1-10', 'P1-F1-11', 'P1-F1-12'],
            'Actividad': [
                'SPC R1 - COORD. REQUISITOS COMERCIALES',
                'SPC R2 - VISITA TÉCNICA PREVIA',
                'SPC R3 - RECEPCIÓN DE DOCUMENTOS',
                'SPC R4 - REVISION DE DOCUMENTOS',
                'SPC R5- CAPACITACION TECNICA',
                'SPC R6 - CATASTRO',
                'SPC R7 - INFORME TECNICO',
                'SPC R8 - CUMPLIMIENTO DE REQUISITOS',
                'SPC R9 - REGISTRO DE SOLICITUDES',
                'SPC R10 - EMISIÓN DE PRESUPUESTOS',
                'SPC R11 - PAGOS',
                'SPC R12 - INSTALACIÓN DE CONEXIONES'
            ],
            'Fase': ['SPC'] * 12,
            'Owner': ['SPC'] * 12,
            'Inicio': ['2025-05-08', '2025-10-06', '2025-09-02', '2025-09-04', '2026-02-04', '2026-03-03',
                      '2026-03-04', '2026-03-02', '2026-03-06', '2026-03-09', '2026-03-11', '2026-03-13'],
            'Fin': ['2025-05-19', '2026-02-08', '2026-03-22', '2026-03-22', '2026-02-16', '2026-03-09',
                   '2026-03-10', '2026-03-31', '2026-03-30', '2026-03-30', '2026-03-31', '2026-04-07'],
            'Estado': ['Terminado', 'Terminado', 'En progreso', 'En progreso', 'Terminado', 'En progreso',
                      'En progreso', 'En progreso', 'No iniciado', 'No iniciado', 'No iniciado', 'No iniciado'],
            'Avance': [1.00, 1.00, 0.7138, 0.7138, 1.00, 0.70, 0.63, 0.069, 0.00, 0.00, 0.00, 0.00],
            'Peso': [0.02, 0.02, 0.02, 0.02, 0.02, 0.02, 0.04, 0.10, 0.02, 0.10, 0.02, 0.60]
        })
        spc_df.to_excel(writer, sheet_name='SPC', index=False)
    
    print(f"✅ Plantilla Excel creada: {output_file}")

if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1:
        if sys.argv[1] == '--crear-plantilla':
            create_template_excel()
        else:
            excel_file = sys.argv[1]
            if os.path.exists(excel_file):
                config, proyectos, actividades = excel_to_js(excel_file)
                print("\n📋 Resumen:")
                print(f"   Fecha corte: {config.get('Fecha Corte', 'No definida')}")
                for p in proyectos:
                    acts = [a for a in actividades if a['owner'] == p['owner']]
                    print(f"   • {p['nombre']}: {len(acts)} actividades, {p['tipo_calculo']}")
            else:
                print(f"❌ Error: No se encuentra {excel_file}")
    else:
        print("Uso:")
        print("  python excel_to_js.py --crear-plantilla  # Crear plantilla Excel")
        print("  python excel_to_js.py proyecto.xlsx      # Generar data.js desde Excel")

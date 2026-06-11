import sys
import os
sys.path.insert(0, os.path.abspath("src"))
import openpyxl
from rddata import DataEngine

engine = DataEngine("Registro_2026.xlsx")
wb = engine._wb_cache
grados = engine.obtener_grados_activos(wb=wb)

print("Grados:", grados)
for g in grados:
    print(f"\n--- Grade {g} ---")
    estudiantes = engine.obtener_estudiantes_completos(g, wb=wb)
    materias_grado = engine.obtener_materias_por_grado(g, wb=wb)
    materias_limpias = [m for m in materias_grado if m and m not in ["Sin materias", "No hay materias", "General", "Todas las Materias"]]
    print("Materias limpias:", materias_limpias)
    
    # Identificar materias de tecnología y materias normales
    tech_mats = []
    norm_mats = []
    for m in materias_limpias:
        m_upper = engine.limpiar_acentos(m)
        if "HOGAR" in m_upper or "DESARROLLO" in m_upper or "AGRO" in m_upper or "COMERCIO" in m_upper:
            tech_mats.append(m)
        else:
            norm_mats.append(m)
            
    print("Tech mats:", tech_mats)
    print("Norm mats:", norm_mats)
    
    # Obtener notas
    notas_por_materia_trimestre = {}
    for trim in ["Trimestre 1", "Trimestre 2", "Trimestre 3"]:
        for m in materias_limpias:
            proms = engine.obtener_promedios_reales(g, m, trim, wb=wb)
            if proms:
                notas_por_materia_trimestre[(m, trim)] = proms
                
    print("Resolved combinations in cache:", list(notas_por_materia_trimestre.keys()))
    
    for est in estudiantes[:5]:
        nom = est["nombre"]
        promedios_trimestrales = []
        for trim in ["Trimestre 1", "Trimestre 2", "Trimestre 3"]:
            norm_notas = []
            for m in norm_mats:
                proms_dict = notas_por_materia_trimestre.get((m, trim), {})
                if nom in proms_dict and proms_dict[nom] is not None:
                    norm_notas.append(proms_dict[nom])
            
            tech_notas = []
            for m in tech_mats:
                proms_dict = notas_por_materia_trimestre.get((m, trim), {})
                if nom in proms_dict and proms_dict[nom] is not None:
                    tech_notas.append(proms_dict[nom])
            
            tec_prom = sum(tech_notas) / len(tech_notas) if tech_notas else None
            trim_notas = list(norm_notas)
            if tec_prom is not None:
                trim_notas.append(tec_prom)
                
            if trim_notas:
                trim_avg = sum(trim_notas) / len(trim_notas)
                promedios_trimestrales.append(trim_avg)
                
        print(f"Student {nom}: promedios_trimestrales={promedios_trimestrales}")

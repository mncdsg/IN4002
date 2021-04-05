import pandas as pd
import matplotlib.pyplot as plt
import xlsxwriter

def elimceros(lista):
    listaaux = []
    for i in range(len(lista)):
        if lista[i] != 0:
            listaaux.append(lista[i])
    return(listaaux)

def elimvalores(lista1, lista2):    
    listaaux = []  
    for i in range(len(lista1)):
        if lista1[i] != 0:
            listaaux.append(lista2[i])
    return(listaaux)

def elimnan(lista):
    listaaux = []
    for i in range(len(lista)):
        if type(lista[i]) == str:
            listaaux.append(lista[i])
    return(listaaux)
            

df = pd.read_excel('Retroalimentación Presentaciones Sección 2 (respuestas).xlsx', 
                    sheet_name='Respuestas de formulario 1')

equipos = ['SYNCOM', 'Niksen', 'Ataraxia', 'Alerce']

comunicacion = ['Vocabulario y gramática', 'Dicción y fluidez','Diseño de láminas',
                'Motivación/ energía']
ind_comunicacion = ['mínimo','acercándose','logrado','fortaleza','no evaluado']

contenido = ['Profundidad de la investigación','Datos, gráficos e info relevante',
             'Justificación y argumentos','Hilo conductor','Relevancia, impacto, tracción']
ind_contenido = ['mínima','falta desarrollar','suficiente','fortaleza','no evaluado']


for equipo in range(len(equipos)):
    
    resultados_com = [0]*4
    for aspecto in range(0,4):
        listaaux1 = [0]*5
        for indicador in range(0,5):
            listaaux1[indicador] = len(df[
                (df['¿A qué grupo estás dando feedback?'] == equipos[equipo]) &
                (df['Comunicación  '+'['+comunicacion[aspecto]+']'] == ind_comunicacion[indicador])
                 ]
                )
        resultados_com[aspecto] = listaaux1
    
    resultados_cont = [0]*5
    for aspecto in range(0,5):
        listaaux2 = [0]*5
        for indicador in range(0,5):
            listaaux2[indicador] = len(df[
                        (df['¿A qué grupo estás dando feedback?'] == equipos[equipo]) &
                        (df['Contenido '+'['+contenido[aspecto]+']'] == ind_contenido[indicador])
                         ]
                        )
        resultados_cont[aspecto] = listaaux2
                    
    plt.rcParams['font.size'] = 5
    fig, axs = plt.subplots(2, 2)
    axs[0, 0].pie(elimceros(resultados_com[0]), labels=elimvalores(resultados_com[0],ind_comunicacion), 
                  autopct="%0.1f %%", explode=[0.1]*len(elimceros(resultados_com[0])))
    axs[0, 0].set_title(comunicacion[0], fontweight='bold')
    axs[0, 1].pie(elimceros(resultados_com[1]), labels=elimvalores(resultados_com[1],ind_comunicacion), 
                  autopct="%0.1f %%", explode=[0.1]*len(elimceros(resultados_com[1])))
    axs[0, 1].set_title(comunicacion[1], fontweight='bold')
    axs[1, 0].pie(elimceros(resultados_com[2]), labels=elimvalores(resultados_com[2],ind_comunicacion), 
                  autopct="%0.1f %%", explode=[0.1]*len(elimceros(resultados_com[2])))
    axs[1, 0].set_title(comunicacion[2], fontweight='bold')
    axs[1, 1].pie(elimceros(resultados_com[3]), labels=elimvalores(resultados_com[3],ind_comunicacion), 
                  autopct="%0.1f %%", explode=[0.1]*len(elimceros(resultados_com[3])))
    axs[1, 1].set_title(comunicacion[3], fontweight='bold')
    fig.suptitle('Resultados Comunicación', size = 16, fontweight='bold')
    fig.tight_layout(pad = 3.0)
    plt.subplots_adjust(top=0.85) 
    fig.savefig('img.png', dpi=1000)
    
    plt.rcParams['font.size'] = 5
    fig2, axs2 = plt.subplots(3, 2)
    axs2[2,1].set_visible(False)
    axs2[0, 0].pie(elimceros(resultados_cont[0]), labels=elimvalores(resultados_cont[0],ind_contenido), 
                  autopct="%0.1f %%", explode=[0.1]*len(elimceros(resultados_cont[0])))
    axs2[0, 0].set_title(contenido[0], fontweight='bold')
    axs2[0, 1].pie(elimceros(resultados_cont[3]), labels=elimvalores(resultados_cont[3],ind_contenido), 
                  autopct="%0.1f %%", explode=[0.1]*len(elimceros(resultados_cont[3])))
    axs2[0, 1].set_title(contenido[3], fontweight='bold')
    axs2[1, 0].pie(elimceros(resultados_cont[2]), labels=elimvalores(resultados_cont[2],ind_contenido), 
                  autopct="%0.1f %%", explode=[0.1]*len(elimceros(resultados_cont[2])))
    axs2[1, 0].set_title(contenido[2], fontweight='bold')
    axs2[1, 1].pie(elimceros(resultados_cont[4]), labels=elimvalores(resultados_cont[4],ind_contenido), 
                  autopct="%0.1f %%", explode=[0.1]*len(elimceros(resultados_cont[4])))
    axs2[1, 1].set_title(contenido[4], fontweight='bold')
    axs2[2, 0].pie(elimceros(resultados_cont[1]), labels=elimvalores(resultados_cont[1],ind_contenido), 
                  autopct="%0.1f %%", explode=[0.1]*len(elimceros(resultados_cont[1])))
    axs2[2, 0].set_title(contenido[1], fontweight='bold')
    fig2.suptitle('Resultados Contenido', size = 16, fontweight='bold')
    fig2.tight_layout(pad = 3.0)
    plt.subplots_adjust(wspace=None, hspace=None)
    plt.subplots_adjust(top=0.86) 
    fig2.savefig('img2.png', dpi = 1000)
    
    workbook = xlsxwriter.Workbook('Resultados_equipo_'+equipos[equipo]+'.xlsx')
    bold = workbook.add_format({'bold': True})
    worksheet1 = workbook.add_worksheet('Resultados Comunicación')
    worksheet1.set_column('A:A', 2)
    worksheet1.set_column('K:K', 2)
    worksheet1.insert_image('B2', 'img.png')
    worksheet1.write('L2', 'Comentarios Comunicación:', bold)
    worksheet1.write_column('L3', 
                         elimnan(df[(df['¿A qué grupo estás dando feedback?'] == equipos[equipo])
                                    ]['Comentarios respecto a comunicación:'].values.tolist()
                                 )
                         )
    worksheet1.write('L16', 'Escala:', bold)
    worksheet1.write_column('L17', ind_comunicacion)
    
    worksheet2 = workbook.add_worksheet('Resultados Contenido')
    worksheet2.set_column('A:A', 2)
    worksheet2.set_column('K:K', 2)
    worksheet2.insert_image('B2', 'img2.png')
    worksheet2.write('L2', 'Comentarios Contenido:', bold)
    worksheet2.write_column('L3', 
                         elimnan(df[(df['¿A qué grupo estás dando feedback?'] == equipos[equipo])
                                    ]['Comentarios respecto al contenido'].values.tolist()
                                 )
                         )
    worksheet2.write('L16', 'Escala:', bold)
    worksheet2.write_column('L17', ind_contenido)
    
    worksheet3 = workbook.add_worksheet('Comentarios Generales')
    worksheet3.set_column('A:A', 2)
    worksheet3.write('B2', '¿ Cuáles son tus preguntas para este equipo?', bold)
    worksheet3.write_column('B3', 
                         elimnan(df[(df['¿A qué grupo estás dando feedback?'] == equipos[equipo])
                                    ]['¿ Cuáles son tus preguntas para este equipo?'].values.tolist()
                                 )
                         )
    worksheet3.write('C2', '¿Tiene alguna idea/sugerencia de cómo mejorar lo presentado?', bold)
    worksheet3.write_column('C3', 
                         elimnan(df[(df['¿A qué grupo estás dando feedback?'] == equipos[equipo])
                                    ]['¿Tiene alguna idea/sugerencia de cómo mejorar lo presentado?'].values.tolist()
                                 )
                         )
    workbook.close()
    



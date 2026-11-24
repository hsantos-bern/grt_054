# pip install pyinstaller

# pyinstaller --noconfirm --onedir --console --hidden-import "C:\Users\hsantos\Documents\testes 202408.02\python\GRT_054\main_GRT_054.py"



import os, sys
from datetime import date, datetime
from configparser import ConfigParser
from time import sleep
import logging

import shutil
import smtplib
import mimetypes
from email.message import EmailMessage

for ia1 in range(1): #~~~~~~~~~~~~~~~~ FUNÇÕES minhas

     
    def aplicaConfig(caminho_arquivo_Config, travas, TravaTeste):

        config = ConfigParser()
        # Abra o arquivo config.ini com a codificação UTF-8


        try:


                with open(caminho_arquivo_Config, 'r', encoding='utf-8') as f:
                    config.read_file(f)
                
                if config.has_option(travas, TravaTeste):

                    saida =  config.get(travas, TravaTeste)
                    
                    return saida
                
                else:
                    logger.warning(f'Chave [{travas}] {TravaTeste} não encontrada em {caminho_arquivo_Config}')
                    return False
            
        except Exception as e:
            #logger.error(f'Erro ao recuperar informação do arquivo Config: {e}')
            return False
     
    #------------------------\\
    
    def configurar_logger(caminho_log):
        # garante que a pasta existe
        os.makedirs(os.path.dirname(caminho_log), exist_ok=True)

        # cria o logger
        logger = logging.getLogger("meu_logger")
        logger.setLevel(logging.INFO)

        # remove handlers antigos (se existirem)
        for h in logger.handlers[:]:
            logger.removeHandler(h)

        # cria handler novo para o arquivo desejado
        handler = logging.FileHandler(caminho_log, encoding="utf-8")
        # formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")

        formatter = logging.Formatter(
        fmt="%(asctime)s - %(levelname)s - %(message)s",
        datefmt="%d.%m.%Y %H:%M:%S"
        )
        
        handler.setFormatter(formatter)

        logger.addHandler(handler)
        return logger

    #------------------------\\    
   
    def send_email(sender_email, sender_password, recipient_email,cc_emails, bcc_emails, subject, body, ext_smtp_server, ext_smtp_port, caminho_arquivo_excel='', caminho_pasta=[]):
        # import os
        # #import time
        # import smtplib
        # import mimetypes
        # from email.message import EmailMessage


        try:

            # Configurando o servidor SMTP
            smtp_server = ext_smtp_server#'grid171.mailgrid.com.br'
            smtp_port = ext_smtp_port#587

            # Inicializando o objeto EmailMessage
            msg = EmailMessage()
            msg['Subject'] = subject
            msg['From'] = sender_email
            
            if recipient_email != '':
                msg['To'] = recipient_email
            
            if cc_emails != '':
                msg['Cc'] = cc_emails

            msg['Bcc'] = bcc_emails
            #msg['Bcc'] = ', '.join(bcc_emails) if bcc_emails else ""

            # --- Configurar Alta Prioridade ---
            msg.add_header('X-Priority', '1 (Highest)') # Alta prioridade
            msg.add_header('Importance', 'High')        # Essencial para Outlook/Gmail

            #msg.set_content(body)
            msg.add_alternative(body, subtype="html")
            #~~~~~~~~~~~~~~~~~~~~~~~~~ \\ ~~~~~~~
            
            # Anexando arquivos ao email, se houver

            # se for passado um único arquivo, lê o arquivo e anexa
            if caminho_arquivo_excel !='':            

                with open(caminho_arquivo_excel, 'rb') as arquivo:
                    tipo_mime, _ = mimetypes.guess_type(caminho_arquivo_excel)
                    if tipo_mime is None:
                        tipo_mime = 'application/octet-stream'
                    tipo, subtipo = tipo_mime.split('/')
                    nome_arquivo1 = caminho_arquivo_excel.split('\\')[-1]
                    msg.add_attachment(arquivo.read(), maintype=tipo, subtype=subtipo, filename=nome_arquivo1)

            #~~~~~~~~~~~~~~~~~~~~~~~~~ \\ ~~~~~~~

            # se for passada uma pasta, lê a pasta e anexa os arquivos
            if caminho_pasta != []:
            
                for file in caminho_pasta:
                    attchment = open(file, 'rb')
                    file_data=attchment.read()
                    file_name = os.path.basename(file)

                    msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)
                
            #~~~~~~~~~~~~~~~~~~~~~~~~~ \\ ~~~~~~~
            # Conectando e autenticando no servidor SMTP
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            server.login(sender_email, sender_password)

            # Enviando o email

            server.send_message(msg)

            #time.sleep(60)
            # Encerrando a conexão com o servidor SMTP
            server.quit()

            return True  # Indica que o email foi enviado com sucesso

        except Exception as e:
            
            fraseerro="Ocorreu um erro durante o envio do email: "+ str(e)
            logger.error(fraseerro)
            
            return False  # Indica que houve um erro ao enviar o email

    #------------------------\\
    

    
#########################################################################


for ia1 in range(1): #~~~~ VARIÁVEIS INTERNAS ~~~~~ dados pastas da automação
        
        
        if getattr(sys, 'frozen', False):

            diretorio_atual = os.path.dirname(sys.executable)
            
        else:
            
            diretorio_atual = os.path.dirname(os.path.abspath(__file__))            
       
        
        caminho_diretorio_arquivos = os.path.join(diretorio_atual, "arquivos_externos")

        #~~~~~~~~~~~~~~~~~ LOG's ~~~~~~~~~~~
        data_e_hora_atuais1 = datetime.now().strftime("%Y%m.%d - %H.%M.%S")
        log_filename = f"{data_e_hora_atuais1}.log"
        caminho_arquivo_Logs  = os.path.join(diretorio_atual, log_filename)

        logger = configurar_logger(caminho_arquivo_Logs)

        ############
       
        #~~~~~~~~~~~~~~~~~ CONFIG ~~~~~~~~~~~
        
        caminho_Config_Basic  = os.path.join(caminho_diretorio_arquivos, "config.ini")

        if not os.path.exists(caminho_Config_Basic):
            logger.error(f'Arquivo "config" não foi localizado.')
            sys.exit()

        config_local = aplicaConfig(caminho_Config_Basic, 'configuracoes1','arquivo_config')
        caminho_arquivo_Config  = os.path.join(caminho_diretorio_arquivos, f"{config_local}.ini")

        if not os.path.exists(caminho_arquivo_Config):
            logger.error(f'Arquivo "config" desta Automação não foi localizado.')
            sys.exit()
        
       
        for i1 in range(1):#~~~~~~~~~~~~~~~~~~~~~~~~~~~ dados servidor SMTP para envio de email
            sender_email = aplicaConfig(caminho_arquivo_Config, 'smtp','login')
            sender_password = aplicaConfig(caminho_arquivo_Config, 'smtp','password')
            sender_smtp_server = aplicaConfig(caminho_arquivo_Config, 'smtp','server')
            sender_smtp_port = aplicaConfig(caminho_arquivo_Config, 'smtp','port')
            logger.info("dados de SMTP carregados.")

        for i1 in range(1):#~~~~~~~~~~~~~~~~~~ Email de Início, Fim e Erro na Automação: pra quem

            recipient_email = aplicaConfig(caminho_arquivo_Config, 'configuracoes1','email_Cliente')
            recipient_email2    = aplicaConfig(caminho_arquivo_Config, 'configuracoes1','email_interno')
            recipient_email_cc  = aplicaConfig(caminho_arquivo_Config, 'configuracoes1','email_cc')
            recipient_email_bcc = aplicaConfig(caminho_arquivo_Config, 'configuracoes1','email_cco')

            logger.info("dados de ENVIO carregados.")
            


        #~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        numeroAutomacao = aplicaConfig(caminho_arquivo_Config, 'configuracoes1','numeroAutomacao')
        descriAutomacao = aplicaConfig(caminho_arquivo_Config, 'configuracoes1','descricaoAutomacao')
        RoboAtual = f'{numeroAutomacao}_{descriAutomacao}'

       
        ############
        logger.info(f"Iniciando a automação: {RoboAtual}.")

        path_downloads_folder1 = os.path.join("C:\\", numeroAutomacao)
        
        if not os.path.exists(path_downloads_folder1):
            os.makedirs(path_downloads_folder1)

        data_hoje = datetime.now().strftime("%d.%m.%Y")
        data_hoje1 = datetime.now().strftime("%Y%m.%d")
        
        path_downloads_folder = os.path.join(path_downloads_folder1, data_hoje1)

        if not os.path.exists(path_downloads_folder):
            os.makedirs(path_downloads_folder)

        finalizado_file1 = os.path.join(path_downloads_folder, f"finalizado_{data_hoje}.txt")        

        if os.path.exists(finalizado_file1):
            logger.info("automação já executada hoje.")
            sys.exit()
            # pass
        
        ############        
       
        #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        for i1 in range(1):#~~~~~~~~~~~~ pastas e arquivos no onedrive

            pasta_Destino = aplicaConfig(caminho_arquivo_Config, 'configuracoes1','arquivo_destino')

            nomePasta1 = 'OP-GRT'
            nomePasta2 = 'usuarios_autom'

            pasta2= os.path.join(nomePasta1, nomePasta2)
            caminho_Fonte= f"{pasta2}\\config.ini"

            pasta3= os.path.join(pasta_Destino, nomePasta1)
            pasta4= os.path.join(pasta3, nomePasta2)

            arquivoFinal= os.path.join(pasta4, 'config.ini')
          

#====================== \\ ====================== \\


for i1 in range(1):#~~~~~~ # Projeto de fato

    erro_Ocorreu = 0

    logger.info(f"Dados carregados.")

    if os.path.exists(arquivoFinal):
        os.remove(arquivoFinal)


    else:    
       
        #fonte
        # pasta1=r'OP-GRT'
        if not os.path.exists(nomePasta1):
                os.makedirs(nomePasta1)
            
        
        if not os.path.exists(pasta2):
                os.makedirs(pasta2)

        if not os.path.exists(caminho_Fonte):

            msgErro=f"O arquivo config que contém os dados não foi localizado em '{pasta2}'."
            erro_Ocorreu +=1
            logger.warning(msgErro)
            print(msgErro)
            break
        #---------------------------------------------

        #destino

        if not os.path.exists(pasta3):
            
            try:
                os.makedirs(pasta3)
                
            except Exception as e:
                
                msgErro=f"Ocorreu um erro ao tentar acessar '{pasta3}'."
                erro_Ocorreu +=1
                logger.warning(msgErro)
                print(msgErro)
                break
        
        
        if not os.path.exists(pasta4):
            
            try:
                os.makedirs(pasta4)
                
            except Exception as e:
                
                msgErro=f"Ocorreu um erro ao tentar acessar '{pasta4}'."
                erro_Ocorreu +=1
                logger.warning(msgErro)
                print(msgErro)
                break


    try:

        shutil.copy2(caminho_Fonte, arquivoFinal)

    except Exception as e:
        
        msgErro=f"Ocorreu um erro ao tentar copiar da pasta fonte pra pasta destino."
        erro_Ocorreu +=1
        logger.warning(msgErro)
        print(msgErro)
        
#=========================\\=========================\\=========================\\

for i1 in range(1):#~~~~~~ # acima é da macro, abaixo é para o agendador
    
    data_hoje = datetime.now().strftime("%d.%m.%Y")

    if erro_Ocorreu == 0 :

        finalizado_file = os.path.join(path_downloads_folder, f"finalizado_{data_hoje}.txt")

    else:
        finalizado_file = os.path.join(path_downloads_folder, f"finalizado_com_excecao_{data_hoje}.txt")
        
        conteudo_arquivoBodyEmail = f"""
            <p>{RoboAtual}</p>
            <p></p>
            <p>Olá,</p>
            <p>Ocorreu uma exceção durante a execução da automação.</p>
            <p>O arquivo de logs está anexado para visualizar o que houve.
            </p>
            <p><i>*Favor não responder este e-mail.</i></p>
            <p>Atenciosamente,<br>
            Bernhoeft</p>
            """

        logger.info("Enviando email de erro.")

        db2 = send_email(sender_email, sender_password, recipient_email_cc, 
                                '', '', "erro de execução", conteudo_arquivoBodyEmail, 
                                sender_smtp_server, sender_smtp_port, caminho_arquivo_Logs)
                                            
        
        if db2 == False:
            logger.error("Problema no envio do email.")
            erro_Ocorreu = erro_Ocorreu + 1



    # Se o arquivo não existir, cria o arquivo
    with open(finalizado_file, 'w', encoding='utf-8') as arquivo:
        pass   


               
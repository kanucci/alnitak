from selenium import webdriver
import time
import aniltakList as aL
import aniltakXlOperator as axOp

class Navegador():

    def __init__(self, login, password):
        self.login = login
        self.password = password
        self.navega_driver_path = 'C:/webdrivers/chromedriver.exe'
        nav_options = webdriver.ChromeOptions()
        nav_options.add_argument('--start-maximized')
        self.driver = webdriver.Chrome(executable_path = self.navega_driver_path, chrome_options = nav_options)

    def conecta(self):
        #driver = webdriver.Chrome(self.navega_driver_path, chrome_options = self.options)
        driver = self.driver
        driver.get("http://hospitalar.postalsaudeservicos.com.br/hospitalarprod/sistema/abertura/login.aspx")

        ologin = driver.find_element_by_id('cphBody_txtLogin')
        osenha = driver.find_element_by_name('ctl00$cphBody$txtSenha')
        obutton = driver.find_element_by_name('ctl00$cphBody$btnSalvar')

        ologin.send_keys(self.login)
        osenha.send_keys(self.password)

        obutton.click()
        time.sleep(7)

        try:
            w_f = driver.window_handles[1]
            driver.switch_to.window(w_f)

            sesmt = driver.find_element_by_id('cphBody_dtPapeis_lbkPapel_2')
            sesmt.click()
            time.sleep(5)
            prontuario = driver.find_element_by_id('cphBody_dtlAplicativos_lbtImagem_0')
            pronturio.click()
            print('Successfully Connection!')
            time.sleep(5)

            #matricula = driver.find_element_by_name('ctl00$cphBody$Conteudo$txtNome')
            #pesquisaPac = driver.find_element_by_name('ctl00$cphBody$Conteudo$btnPesquisaPaciente')
            #matricula.send_keys('84133678')
            plan_list = aL.xl_List()
            lista_aso = []

            for plan in plan_list:

                xlPlan = axOp.oPlan(plan, lista_aso)
                lista_aso = xlPlan.operator()

            #print(lista_aso)
            #print(lista_aso[2][1])
            l_a = len(lista_aso)-1

            for aso in range(l_a):
                #print(aso)
                #print(lista_aso[aso][1])
                matricula = driver.find_element_by_name('ctl00$cphBody$Conteudo$txtNome')
                pesquisaPac = driver.find_element_by_name('ctl00$cphBody$Conteudo$btnPesquisaPaciente')
                matricula.send_keys(lista_aso[aso][1])
                pesquisaPac.click()
                time.sleep(8)

                carteirinhas = len(driver.find_elements_by_xpath('//*[@id="cphBody_Conteudo_grdPessoa"]/tbody/tr/td[7]'))

                if carteirinhas != 1:
                    for carteirinha in range(carteirinhas):
                        item = driver.find_elements_by_xpath('//*[@id="cphBody_Conteudo_grdPessoa"]/tbody/tr/td[7]')[carteirinha].text
                        #flag += 1
                        #print(item[9:11])
                        if item[9:11] == '00':
                            driver.find_elements_by_xpath('//*[@id="cphBody_Conteudo_grdPessoa"]/tbody/tr/td[7]')[carteirinha].click()
                            time.sleep(5)

                if driver.find_element_by_name('ctl00$cphBody$Conteudo$btnAplica'):
                    confirmar = driver.find_element_by_name('ctl00$cphBody$Conteudo$btnAplica')
                    confirmar.click()
                    time.sleep(4)

                pesquisa = driver.find_element_by_name('ctl00$cphBody$Conteudo$btnPesquisa')
                pesquisa.click()
                time.sleep(4)

                asos_data = []
                asos_tipo = []
                tipo = 0
                data = 0
                #asos = driver.find_elements_by_xpath('//*[@id="cphBody_Conteudo_grdAtendimentos"]/tbody/tr/td[3]')
                asos_data = len(driver.find_elements_by_xpath('//*[@id="cphBody_Conteudo_grdAtendimentos"]/tbody/tr/td[2]'))
                asos_tipo = len(driver.find_elements_by_xpath('//*[@id="cphBody_Conteudo_grdAtendimentos"]/tbody/tr/td[3]'))

                #print(tipo)
                #print(asos_tipo)

                for tipo in range(asos_tipo):
                    for data in range(asos_data):
                        if driver.find_elements_by_xpath('//*[@id="cphBody_Conteudo_grdAtendimentos"]/tbody/tr/td[3]')[tipo].text == lista_aso[aso][3] and driver.find_elements_by_xpath('//*[@id="cphBody_Conteudo_grdAtendimentos"]/tbody/tr/td[2]')[data].text == lista_aso[aso][2]:
                            driver.find_elements_by_xpath('//*[@id="cphBody_Conteudo_grdAtendimentos"]/tbody/tr/td[2]')[tipo].click()
                            time.sleep(3)
                            btArquivos = driver.find_element_by_id('cphBody_Conteudo_btnArquivos')
                            btArquivos.click()
                            time.sleep(10)
                            try:
                                #arquivo = driver.find_element_by_id('cphBody_Conteudo_grdItens')
                                #arquivo.click()
                                print('Arquivo selecionado')
                                #time.sleep(20)
                                btDownload = driver.find_element_by_id('cphBody_Conteudo_btnBaixar')
                                btDownload.click()
                                time.sleep(40)
                                w_s = driver.window_handles[2]
                                driver.switch_to.window(w_s)
                                driver.execute_script("window.close('');")
                                driver.switch_to.window(w_f)
                                lista_aso[aso].append('ok')
                                time.sleep(4)

                            except:

                                print('Passei no except, cuidado!')
                                btVoltar = driver.find_element_by_id('cphBody_lbtVoltar')
                                btVoltar.click()
                                time.sleep(4)
                                pass

                            break

                btVoltar = driver.find_element_by_id('cphBody_lbtVoltar')
                btVoltar.click()
                time.sleep(4)
                pesquisaPac = driver.find_element_by_name('ctl00$cphBody$Conteudo$btnPesquisaPaciente')
                pesquisaPac.click()
                time.sleep(3)
                #print(lista_aso[aso][1])

            for plan in plan_list:

                xlPlan = axOp.oPlan(plan, lista_aso)
                xlPlan.updateXl()  

        except Exception as e:

            login_status = driver.find_element_by_id('cphBody_divErro')
            print(login_status.text)
            driver.quit()


    def encerrar_conexao(self):
        self.driver.quit()

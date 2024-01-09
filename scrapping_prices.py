import sys
import time
import pandas as pd
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager

class FileDialogApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Seleção de Planilha e Salvamento')
        self.setGeometry(300, 300, 300, 200)

        layout = QVBoxLayout()

        self.openFileNameButton = QPushButton('Abrir Planilha Excel', self)
        self.openFileNameButton.clicked.connect(self.openFileNameDialog)
        layout.addWidget(self.openFileNameButton)

        self.saveFileNameButton = QPushButton('Salvar Planilha Excel', self)
        self.saveFileNameButton.clicked.connect(self.saveFileDialog)
        layout.addWidget(self.saveFileNameButton)

        self.pathLabel = QLabel(self)
        layout.addWidget(self.pathLabel)

        self.setLayout(layout)

    def openFileNameDialog(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self, "QFileDialog.getOpenFileName()", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
        if fileName:
            self.pathLabel.setText(f"Arquivo Aberto: {fileName}")
            self.read_and_process_excel(fileName)

    def saveFileDialog(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getSaveFileName(self, "QFileDialog.getSaveFileName()", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
        if fileName:
            self.save_excel(fileName)

    def read_and_process_excel(self, file_path):
        self.df = pd.read_excel(file_path)
        self.df['OBRAMAX'] = self.df['OBRAMAX'].apply(self.scrape_obramax)
        self.df['LEROY'] = self.df['LEROY'].apply(self.scrape_leroy)
        self.df['JOLI'] = self.df['JOLI'].apply(self.scrape_joli)
        self.df['COPAFER'] = self.df['COPAFER'].apply(self.scrape_copafer)

    def scrape_obramax(self, url):
        if pd.isna(url) or url == "":
            return "Link não encontrado"

        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)

        try:
            driver.get(url)
            time.sleep(5)  # Esperar o conteúdo ser carregado

            try:
                # Primeiro, tentar encontrar o preço promocional
                price_element = driver.find_element(By.CLASS_NAME, 'lojaobramax-store-components-0-x-currencyContainer')
                if price_element:
                    price = price_element.text.strip()
                    return price
            except NoSuchElementException:
                # Se o preço promocional não estiver presente, buscar o preço padrão
                price_element = driver.find_element(By.CLASS_NAME, 'vtex-product-price-1-x-currencyContainer')
                if price_element:
                    price = price_element.text.strip()
                    return price

            return 'Preço não encontrado'
        except Exception as e:
            return f'Erro: {e}'
        finally:
            driver.quit()

    def scrape_leroy(self, url):
        if pd.isna(url) or url == "":
            return "Link não encontrado"

        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)
        try:
            driver.get(url)
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'span.css-rwb0cd-to-price__integer.e17u5sne6'))
            )
            price_element = driver.find_element(By.CSS_SELECTOR, 'span.css-rwb0cd-to-price__integer.e17u5sne6')
            price = price_element.text.strip()
            return price
        except Exception as e:
            return f'Erro: {e}'
        finally:
            driver.quit()

    def scrape_joli(self, url):
        if pd.isna(url) or url == "":
            return "Link não encontrado"
        
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)
        try:
            driver.get(url)
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'joli-joli-theme-0-x-price'))
            )
            price_element = driver.find_element(By.CLASS_NAME, 'joli-joli-theme-0-x-price')
            price = price_element.text.strip()
            return price
        except Exception as e:
            return f'Erro: {e}'
        finally:
            driver.quit()

    # O método scrape_copafer ainda precisa da lógica correta para o scraping
    def scrape_copafer(self, url):
        if pd.isna(url) or url == "":
            return "Link não encontrado"
        
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)
        try:
            driver.get(url)
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'preco'))
            )
            price_element = driver.find_element(By.CLASS_NAME, 'preco')
            price = price_element.text.strip()
            return price
        except Exception as e:
            return f'Erro: {e}'
        finally:
            driver.quit()

    def save_excel(self, file_path):
        self.df.to_excel(file_path, index=False)
        self.pathLabel.setText(f"Arquivo Salvo: {file_path}")

def main():
    app = QApplication(sys.argv)
    ex = FileDialogApp()
    ex.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()

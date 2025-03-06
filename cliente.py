import sys
import os
import socket
import json
import platform
import getpass
from datetime import datetime

# PyQt
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton, QLabel, QMessageBox
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QGuiApplication

def add_to_startup():
    """
    Cria um arquivo .bat no diretório de inicialização do Windows, 
    para executar este script automaticamente ao iniciar o sistema.
    """
    # Caminho para a pasta Startup do usuário
    startup_folder = os.path.join(os.getenv('APPDATA'), 
                                  r"Microsoft\Windows\Start Menu\Programs\Startup")

    # Arquivo .bat que vamos criar
    bat_path = os.path.join(startup_folder, "registro_ponto.bat")

    # Se já existir, não criamos de novo (ou pode sobrescrever se preferir)
    if os.path.exists(bat_path):
        return  # Já existe, não faz nada

    # Caminho do Python que está rodando este script
    python_exe = sys.executable
    # Caminho deste arquivo .py
    script_path = os.path.abspath(__file__)

    # Conteúdo do .bat para rodar o script
    # Obs: "start" evitaria que o terminal ficasse aberto, mas pode abrir janelas extras.
    # Aqui deixamos simples, chamando diretamente:
    bat_content = f'@echo off\n"{python_exe}" "{script_path}"\n'

    # Cria o .bat no diretório de Startup
    with open(bat_path, 'w', encoding='utf-8') as f:
        f.write(bat_content)

def create_main_window():
    """
    Cria e retorna um QMainWindow em modo kiosk (tela cheia, sem barra de título),
    com layout simples para registrar ponto.
    """
    window = QMainWindow()
    # Remove a barra de título/decoração
    window.setWindowFlags(Qt.FramelessWindowHint)

    central_widget = QWidget()
    window.setCentralWidget(central_widget)
    layout = QVBoxLayout(central_widget)
    layout.setAlignment(Qt.AlignCenter)

    # Título em fonte grande
    label_titulo = QLabel("Sistema de Registro de Ponto")
    label_titulo.setStyleSheet("""
        font-size: 48px;
        font-weight: bold;
        color: #333;
    """)
    label_titulo.setAlignment(Qt.AlignCenter)
    layout.addWidget(label_titulo)

    # Botão "Registrar Ponto"
    btn_registrar = QPushButton("Registrar Ponto")
    btn_registrar.setStyleSheet("""
        font-size: 36px;
        padding: 20px;
        background-color: #007ACC;
        color: white;
        border-radius: 10px;
    """)
    layout.addWidget(btn_registrar)

    # Label de status
    label_status = QLabel("Clique para registrar seu ponto.")
    label_status.setStyleSheet("font-size: 24px; color: #555;")
    label_status.setAlignment(Qt.AlignCenter)
    layout.addWidget(label_status)

    # Função chamada ao clicar no botão
    def on_registrar_ponto():
        username = getpass.getuser()
        windows_info = platform.platform()
        timestamp = datetime.now().isoformat()

        data = {
            "username": username,
            "windows_info": windows_info,
            "timestamp": timestamp
        }

        try:
            # Ajuste aqui o IP/porta do seu servidor
            host = '192.168.51.8'
            port = 5000

            client_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            client_socket.connect((host, port))
            client_socket.send(json.dumps(data).encode("utf-8"))

            response_data = client_socket.recv(1024).decode("utf-8")
            response = json.loads(response_data)
            client_socket.close()

            if response.get("status") == "success":
                hora_registro = datetime.now().strftime("%H:%M:%S")
                label_status.setText(f"Ponto registrado hoje às {hora_registro}.")
                QMessageBox.information(window, "Sucesso", response.get("message", "Ponto registrado!"))
            else:
                msg = response.get("message", "Erro ao registrar ponto.")
                QMessageBox.warning(window, "Erro", msg)
        except Exception as e:
            QMessageBox.critical(window, "Erro", f"Falha na conexão com o servidor:\n{str(e)}")

    btn_registrar.clicked.connect(on_registrar_ponto)

    return window

def main():
    # 1) Cria o .bat no Startup do Windows para rodar automaticamente no login (se ainda não existir).
    add_to_startup()

    # 2) Inicia a aplicação PyQt
    app = QApplication(sys.argv)

    # 3) Cria uma janela para cada tela (monitor) e exibe em tela cheia
    windows = []
    screens = QGuiApplication.screens()
    if not screens:
        print("Nenhuma tela detectada!")
        sys.exit(1)

    for screen in screens:
        w = create_main_window()
        # Move a janela para a área do monitor específico
        geometry = screen.geometry()
        w.setGeometry(geometry)
        # Exibe em fullScreen
        w.showFullScreen()
        windows.append(w)

    sys.exit(app.exec_())

if __name__ == "__main__":
    main()

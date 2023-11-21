from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.textinput import TextInput
import pandas as pd
from plyer import sms

class ExcelApp(App):
    def build(self):
        # Layout principal
        layout = BoxLayout(orientation='vertical', spacing=10, padding=10)

        # Botão para carregar planilha
        load_button = Button(text='Carregar Planilha', on_press=self.load_excel, size_hint=(1, 0.1))
        layout.add_widget(load_button)

        # Entrada de mensagem
        self.message_input = TextInput(hint_text='Insira sua mensagem...', size_hint=(1, 0.2))
        layout.add_widget(self.message_input)

        # Botão para enviar SMS
        send_button = Button(text='Enviar SMS para Todos', on_press=self.send_sms_all, size_hint=(1, 0.1))
        layout.add_widget(send_button)

        # Área para exibir informações da planilha
        self.info_label = Label(text='', size_hint=(1, 0.6))
        layout.add_widget(self.info_label)

        return layout

    def load_excel(self, instance):
        # Abrir o FileChooser para selecionar uma planilha
        file_chooser = FileChooserListView()
        file_chooser.bind(on_submit=self.read_excel)
        self.root.add_widget(file_chooser)

    def read_excel(self, chooser, file_path, touch):
        # Remover o FileChooser
        self.root.remove_widget(chooser)

        # Carregar a planilha usando pandas
        try:
            self.df = pd.read_excel(file_path[0])
            self.info_label.text = f'Planilha Carregada com Sucesso:\n\n{self.df.head()}'
        except Exception as e:
            # Exibir mensagem de erro se houver algum problema
            self.info_label.text = f'Erro ao Carregar a Planilha:\n\n{str(e)}'

    def send_sms_all(self, instance):
        message = self.message_input.text

        if message and hasattr(self, 'df'):
            for phone_number in self.df['Numero']:
                try:
                    sms.send(recipient=str(phone_number), message=message)
                    print(f"SMS enviado para {phone_number} com sucesso!")
                except NotImplementedError:
                    print("Envio de SMS não suportado nesta plataforma.")
            print("Todos os SMS foram enviados.")
        else:
            print("Por favor, insira a mensagem e carregue a planilha.")

if __name__ == '__main__':
    ExcelApp().run()

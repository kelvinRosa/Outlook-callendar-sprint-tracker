import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, timedelta, date
import requests
from icalendar import Calendar
from dateutil.rrule import rrulestr
from pytz import UTC, timezone
import os
import json
from tkcalendar import DateEntry
import csv
from tkinter.filedialog import asksaveasfilename
from tkinter.filedialog import askopenfilename
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

class CalendarTrackerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Tracker de Sprint")
        
        # Configurações
        self.config_file = "config.json"
        self.excess_minutes_file = "excess_minutes.json"
        self.local_events_file = "local_events.json"
        self.tz_brasil = timezone('America/Sao_Paulo')
        
        # Variáveis
        self.eventos_atuais = []
        self.check_vars = []
        self.excess_minutes = {}  # Dicionário para armazenar minutos excedentes
        self.local_events = []    # Lista para armazenar eventos locais
        
        # Interface
        self.setup_ui()
        self.carregar_config()
        self.carregar_minutos_excedentes()
        self.carregar_eventos_locais()

    def importar_arquivo_ics(self):
        filepath = askopenfilename(
            filetypes=[("Arquivo ICS", "*.ics"), ("Todos os arquivos", "*.*")],
            title="Selecionar arquivo ICS"
        )

        if not filepath:
            return  # usuário cancelou

        try:
            with open(filepath, 'rb') as f:
                ics_data = f.read()

            calendario = Calendar.from_ical(ics_data)

            if self.two_weeks_var.get():
                data_inicio = self.date_picker.get_date()
                data_fim = data_inicio + timedelta(days=13)
            else:
                data_inicio = self.date_picker.get_date()
                data_fim = self.end_date_picker.get_date()

            # Limpar interface
            for widget in self.scrollable_frame.winfo_children():
                widget.destroy()

            self.check_vars = []
            self.eventos_atuais = []

            eventos_calendario = self.processar_calendario(calendario, data_inicio, data_fim)
            eventos_calendario = self.aplicar_filtros(eventos_calendario)
            self.eventos_atuais.extend(eventos_calendario)

            # Carregar eventos locais dentro do período
            eventos_locais_periodo = [
                e for e in self.local_events if data_inicio <= e['start'].date() <= data_fim
            ]

            self.salvar_config()

            for inicio, fim, descricao in sorted(eventos_calendario, key=lambda x: x[0]):
                inicio_br = inicio.astimezone(self.tz_brasil)
                event_key = f"{inicio_br.strftime('%Y%m%d%H%M')}_{descricao}"
                excess_min = self.excess_minutes.get(event_key, 0)
                self.adicionar_evento_na_interface(inicio, fim, descricao, False, excess_min)

            for event in sorted(eventos_locais_periodo, key=lambda x: x['start']):
                self.adicionar_evento_na_interface(
                    event['start'],
                    event['end'],
                    event['description'],
                    True,
                    event.get('excess_minutes', 0)
                )

            self.calcular_total()
            self.atualizar_tarefas_visiveis()

            messagebox.showinfo("Sucesso", "Arquivo ICS importado e eventos carregados.")

        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao importar arquivo ICS:\n{str(e)}")

    def carregar_minutos_excedentes(self):
        """Carrega os minutos excedentes salvos anteriormente"""
        if os.path.exists(self.excess_minutes_file):
            try:
                with open(self.excess_minutes_file, 'r') as f:
                    self.excess_minutes = json.load(f)
            except Exception as e:
                print(f"Erro ao carregar minutos excedentes: {e}")
                self.excess_minutes = {}

    def salvar_minutos_excedentes(self):
        """Salva os minutos excedentes no arquivo"""
        try:
            with open(self.excess_minutes_file, 'w') as f:
                json.dump(self.excess_minutes, f)
        except Exception as e:
            print(f"Erro ao salvar minutos excedentes: {e}")

    def carregar_eventos_locais(self):
        """Carrega os eventos locais salvos anteriormente"""
        if os.path.exists(self.local_events_file):
            try:
                with open(self.local_events_file, 'r') as f:
                    self.local_events = json.load(f)
                    # Converter strings de data para objetos datetime
                    for event in self.local_events:
                        event['start'] = datetime.strptime(event['start'], '%Y-%m-%d %H:%M')
                        event['end'] = datetime.strptime(event['end'], '%Y-%m-%d %H:%M')
            except Exception as e:
                print(f"Erro ao carregar eventos locais: {e}")
                self.local_events = []

    def salvar_eventos_locais(self):
        """Salva os eventos locais no arquivo"""
        try:
            # Converter objetos datetime para strings antes de salvar
            events_to_save = []
            for event in self.local_events:
                event_copy = event.copy()
                event_copy['start'] = event['start'].strftime('%Y-%m-%d %H:%M')
                event_copy['end'] = event['end'].strftime('%Y-%m-%d %H:%M')
                events_to_save.append(event_copy)
                
            with open(self.local_events_file, 'w') as f:
                json.dump(events_to_save, f)
        except Exception as e:
            print(f"Erro ao salvar eventos locais: {e}")

    def adicionar_evento_local(self):
        """Abre uma janela para adicionar um novo evento local"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Adicionar Evento Local")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Campos do formulário
        ttk.Label(dialog, text="Descrição:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        desc_entry = ttk.Entry(dialog, width=40)
        desc_entry.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(dialog, text="Data e Hora de Início:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        start_entry = DateEntry(dialog, date_pattern='yyyy-mm-dd')
        start_entry.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(dialog, text="Hora (HH:MM):").grid(row=1, column=2, padx=5, pady=5, sticky=tk.W)
        start_time = ttk.Entry(dialog, width=5)
        start_time.insert(0, "09:00")
        start_time.grid(row=1, column=3, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(dialog, text="Data e Hora de Fim:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        end_entry = DateEntry(dialog, date_pattern='yyyy-mm-dd')
        end_entry.grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(dialog, text="Hora (HH:MM):").grid(row=2, column=2, padx=5, pady=5, sticky=tk.W)
        end_time = ttk.Entry(dialog, width=5)
        end_time.insert(0, "10:00")
        end_time.grid(row=2, column=3, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(dialog, text="Minutos Excedentes:").grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
        excess_spin = ttk.Spinbox(dialog, from_=0, to=999, width=5)
        excess_spin.set(0)
        excess_spin.grid(row=3, column=1, padx=5, pady=5, sticky=tk.W)
        
        def salvar_evento():
            """Valida e salva o novo evento local"""
            try:
                descricao = desc_entry.get().strip()
                if not descricao:
                    messagebox.showwarning("Aviso", "Por favor, insira uma descrição para o evento")
                    return
                
                # Processar data e hora de início
                start_date = start_entry.get_date()
                start_time_str = start_time.get()
                start_hour, start_minute = map(int, start_time_str.split(':'))
                start_dt = datetime.combine(start_date, datetime.min.time()).replace(
                    hour=start_hour, minute=start_minute)
                
                # Processar data e hora de fim
                end_date = end_entry.get_date()
                end_time_str = end_time.get()
                end_hour, end_minute = map(int, end_time_str.split(':'))
                end_dt = datetime.combine(end_date, datetime.min.time()).replace(
                    hour=end_hour, minute=end_minute)
                
                if end_dt <= start_dt:
                    messagebox.showwarning("Aviso", "A data/hora de fim deve ser após a data/hora de início")
                    return
                
                excess_min = int(excess_spin.get())
                
                # Adicionar o evento à lista
                new_event = {
                    'description': descricao,
                    'start': start_dt,
                    'end': end_dt,
                    'excess_minutes': excess_min
                }
                
                self.local_events.append(new_event)
                self.salvar_eventos_locais()
                
                # Se a data do evento estiver dentro da sprint atual, recarregar os eventos
                current_start = self.date_picker.get_date()
                current_end = current_start + timedelta(days=13)
                
                if current_start <= start_dt.date() <= current_end:
                    self.carregar_eventos()
                
                messagebox.showinfo("Sucesso", "Evento local adicionado com sucesso")
                dialog.destroy()
            
            except ValueError as e:
                messagebox.showerror("Erro", f"Formato de hora inválido. Use HH:MM\n{str(e)}")
            except Exception as e:
                messagebox.showerror("Erro", f"Falha ao adicionar evento:\n{str(e)}")
        
        # Botões
        ttk.Button(dialog, text="Salvar", command=salvar_evento).grid(row=4, column=1, padx=5, pady=10, sticky=tk.E)
        ttk.Button(dialog, text="Cancelar", command=dialog.destroy).grid(row=4, column=2, padx=5, pady=10, sticky=tk.W)

    

    def exportar_csv(self):
        if not self.eventos_atuais and not self.local_events:
            messagebox.showwarning("Aviso", "Não há eventos para exportar")
            return

        filepath = asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("Arquivo CSV", "*.csv"), ("Todos os arquivos", "*.*")],
            title="Salvar como"
        )

        if not filepath:
            return

        try:
            with open(filepath, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile, delimiter=';')

                # --------- Parte 1: Eventos ---------
                writer.writerow([
                    "Data", 
                    "Hora Início", 
                    "Hora Fim", 
                    "Duração (minutos)", 
                    "Excedidos",
                    "Duração Total",
                    "Descrição",
                    "Tipo (Local/Calendário)"
                ])

                eventos_selecionados = []

                for frame in self.scrollable_frame.winfo_children():
                    if hasattr(frame, 'event_data'):
                        for widget in frame.winfo_children():
                            if isinstance(widget, ttk.Checkbutton) and frame.check_var.get() == 1:
                                inicio, fim, descricao = frame.event_data
                                excess_min = int(frame.excess_spin.get())
                                eventos_selecionados.append((inicio, fim, descricao, excess_min, False))
                                break
                    elif hasattr(frame, 'local_event_data'):
                        for widget in frame.winfo_children():
                            if isinstance(widget, ttk.Checkbutton) and frame.check_var.get() == 1:
                                event_data = frame.local_event_data
                                eventos_selecionados.append((
                                    event_data['start'],
                                    event_data['end'],
                                    event_data['description'],
                                    event_data['excess_minutes'],
                                    True
                                ))
                                break

                eventos_selecionados.sort(key=lambda x: self.to_naive_local(x[0]))                
                for inicio, fim, descricao, excess_min, is_local in eventos_selecionados:
                    print(f"[DEBUG] {descricao} - inicio: {inicio} ({inicio.tzinfo}), fim: {fim} ({fim.tzinfo})")
                    inicio_naive = self.to_naive_local(inicio)
                    fim_naive = self.to_naive_local(fim)

                    duracao = int((fim_naive - inicio_naive).total_seconds() / 60)
                    duracao_total = duracao + excess_min

                    writer.writerow([
                        inicio_naive.strftime('%Y-%m-%d'),
                        inicio_naive.strftime('%H:%M'),
                        fim_naive.strftime('%H:%M'),
                        str(duracao),
                        str(excess_min),
                        str(duracao_total),
                        descricao,
                        "Local" if is_local else "Calendário"
                    ])

                # --------- Linha em branco + cabeçalho de tarefas ---------
                writer.writerow([])
                writer.writerow(["Entregas da Sprint"])
                writer.writerow(["Data", "Tarefa"])

                data_inicio = self.date_picker.get_date()
                data_fim = data_inicio + timedelta(days=13) if self.two_weeks_var.get() else self.end_date_picker.get_date()


                for tarefa, data_str in self.tarefas_data.items():
                    try:
                        data = datetime.strptime(data_str, "%Y-%m-%d").date()
                        if data_inicio <= data <= data_fim:
                            writer.writerow([data_str, tarefa])
                    except Exception as e:
                        print(f"Erro ao exportar tarefa '{tarefa}': {e}")

            messagebox.showinfo("Sucesso", f"Eventos e entregas exportados com sucesso para:\n{filepath}")

        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao exportar:\n{str(e)}")

    def to_naive_local(self, dt):
        if isinstance(dt, datetime):
            if dt.tzinfo is not None:
                return dt.astimezone(self.tz_brasil).replace(tzinfo=None)
            return dt
        elif isinstance(dt, date):
            return datetime.combine(dt, datetime.min.time())
        else:
            raise TypeError(f"Esperado datetime/date, recebido: {type(dt)}")
    
    def toggle_end_date(self):
        if self.two_weeks_var.get():
            self.end_date_label.grid_remove()
            self.end_date_picker.grid_remove()
        else:
            self.end_date_label.grid()
            self.end_date_picker.grid()

    def atualizar_tarefas_visiveis(self):
        self.task_listbox.delete(0, tk.END)
        if self.two_weeks_var.get():
            data_inicio = self.date_picker.get_date()
            data_fim = data_inicio + timedelta(days=13)
        else:
            data_inicio = self.date_picker.get_date()
            data_fim = self.end_date_picker.get_date()

        for tarefa, data_str in self.tarefas_data.items():
            try:
                data_tarefa = datetime.strptime(data_str, '%Y-%m-%d').date()
                if data_inicio <= data_tarefa <= data_fim:
                    self.task_listbox.insert(tk.END, tarefa)
            except Exception as e:
                print(f"Erro ao verificar tarefa '{tarefa}': {e}")

    def adicionar_tarefa(self):
        tarefa = self.task_entry.get().strip()
        if tarefa:
            hoje_str = date.today().isoformat()
            self.tarefas_data[tarefa] = hoje_str
            self.task_listbox.insert(tk.END, tarefa)
            self.task_entry.delete(0, tk.END)
            self.salvar_tarefas()

    def salvar_tarefas(self):
        tarefas = []
        for i in range(self.task_listbox.size()):
            texto = self.task_listbox.get(i)
            data = self.tarefas_data.get(texto, date.today().isoformat())
            tarefas.append({"text": texto, "date": data})
        with open(self.task_file, 'w', encoding='utf-8') as f:
            json.dump(tarefas, f, ensure_ascii=False, indent=2)

    def carregar_tarefas(self):
        self.tarefas_data = {}
        if os.path.exists(self.task_file):
            try:
                with open(self.task_file, 'r', encoding='utf-8') as f:
                    tarefas = json.load(f)
                    for t in tarefas:
                        self.tarefas_data[t["text"]] = t["date"]
            except Exception as e:
                print(f"Erro ao carregar entregas: {e}")


    def setup_ui(self):
        self.tabs = ttk.Notebook(self.root)
        self.tabs.pack(fill='both', expand=True)

        # Aba principal (eventos)
        self.main_tab = ttk.Frame(self.tabs)
        self.tabs.add(self.main_tab, text="Meetings")

        mainframe = ttk.Frame(self.main_tab, padding="10")
        mainframe.pack(fill=tk.BOTH, expand=True)

        # Controles superiores
        ttk.Label(mainframe, text="URL do calendário (.ics):").grid(row=0, column=0, sticky=tk.W)
        self.url_entry = ttk.Entry(mainframe, width=50)
        self.url_entry.grid(row=0, column=1, columnspan=4, sticky=(tk.W, tk.E), padx=5, pady=5)
        

        ttk.Label(mainframe, text="Data de início da sprint:").grid(row=1, column=0, sticky=tk.W)
        self.date_picker = DateEntry(mainframe, date_pattern='yyyy-mm-dd')
        self.date_picker.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)

        self.two_weeks_var = tk.BooleanVar(value=True)
        self.two_weeks_check = ttk.Checkbutton(mainframe, text="Sprint de 2 Semanas", variable=self.two_weeks_var, command=self.toggle_end_date)
        self.two_weeks_check.grid(row=1, column=2, sticky=tk.W, padx=5)

        self.end_date_label = ttk.Label(mainframe, text="Data de fim da sprint:")
        self.end_date_picker = DateEntry(mainframe, date_pattern='yyyy-mm-dd')
        self.end_date_label.grid(row=1, column=3, sticky=tk.W)
        self.end_date_picker.grid(row=1, column=4, sticky=tk.W)
        self.toggle_end_date()

        # Botões
        button_frame = ttk.Frame(mainframe)
        button_frame.grid(row=2, column=0, columnspan=5, pady=10, sticky=tk.W)

        ttk.Button(button_frame, text="Carregar de arquivo ICS", command=self.importar_arquivo_ics).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Carregar dar Url", command=self.carregar_eventos).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="+ Adicionar Evento Local", command=self.adicionar_evento_local).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Exportar CSV", command=self.exportar_csv).pack(side=tk.LEFT, padx=5)

        # Resultado
        self.resultado_label = ttk.Label(mainframe, text="Tempo total: 0h00m | Minutos excedentes: 0 | Eventos selecionados: 0")
        self.resultado_label.grid(row=3, column=0, columnspan=5, pady=5)

        # Lista de eventos com scroll
        self.container = ttk.Frame(mainframe)
        self.container.grid(row=4, column=0, columnspan=5, sticky='nsew')

        self.canvas = tk.Canvas(self.container, borderwidth=0, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self.container, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # Configurar pesos da grade
        mainframe.columnconfigure(1, weight=1)
        mainframe.rowconfigure(4, weight=1)

        # Aba de tarefas
        self.task_tab = ttk.Frame(self.tabs)
        self.tabs.add(self.task_tab, text="Entregas da Sprint")

        self.task_entry = ttk.Entry(self.task_tab, width=60)
        self.task_entry.pack(padx=10, pady=(10, 0), anchor="w")

        self.task_add_btn = ttk.Button(self.task_tab, text="Adicionar Entrega", command=self.adicionar_tarefa)
        self.task_add_btn.pack(padx=10, pady=5, anchor="w")

        self.task_listbox = tk.Listbox(self.task_tab, width=80, height=15)
        self.task_listbox.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        self.task_file = "sprint_tasks.json"
        self.carregar_tarefas()

    
    def carregar_config(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                    self.url_entry.insert(0, config.get('url', ''))
                    if 'last_date' in config:
                        self.date_picker.set_date(datetime.strptime(config['last_date'], '%Y-%m-%d').date())
            except Exception as e:
                print(f"Erro ao carregar configuração: {e}")
    
    def salvar_config(self):
        try:
            config = {
                'url': self.url_entry.get(),
                'last_date': self.date_picker.get_date().strftime('%Y-%m-%d')
            }
            with open(self.config_file, 'w') as f:
                json.dump(config, f)
        except Exception as e:
            print(f"Erro ao salvar configuração: {e}")
    
    def carregar_eventos(self):
        url = self.url_entry.get().strip()
        try:
            data_inicio = self.date_picker.get_date()
            if self.two_weeks_var.get():
                data_fim = data_inicio + timedelta(days=13)
            else:
                data_fim = self.end_date_picker.get_date()
        except Exception as e:
            messagebox.showerror("Erro", f"Data inválida: {str(e)}")
            return
        
        # Limpar frame de eventos
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        
        self.check_vars = []
        self.eventos_atuais = []
        
        # Carregar eventos do calendário se houver URL
        if url:
            try:
                if not url.startswith(('http://', 'https://')):
                    url = 'https://' + url
                    
                response = requests.get(url, timeout=10)
                response.raise_for_status()
                calendario = Calendar.from_ical(response.content)
                eventos_calendario = self.processar_calendario(calendario, data_inicio, data_fim)
                eventos_calendario = self.aplicar_filtros(eventos_calendario)
                self.eventos_atuais.extend(eventos_calendario)
            except Exception as e:
                messagebox.showerror("Erro", f"Falha ao carregar calendário: {str(e)}")
        
        # Carregar eventos locais dentro do período
        eventos_locais_periodo = []
        for event in self.local_events:
            start_date = event['start'].date()
            if data_inicio <= start_date <= data_fim:
                eventos_locais_periodo.append(event)
        
        self.salvar_config()
        
        # Adicionar eventos do calendário
        for inicio, fim, descricao in sorted(self.eventos_atuais, key=lambda x: x[0]):
            # Criar chave única para o evento
            inicio_br = inicio.astimezone(self.tz_brasil)
            event_key = f"{inicio_br.strftime('%Y%m%d%H%M')}_{descricao}"
            
            # Obter minutos excedentes salvos
            excess_min = self.excess_minutes.get(event_key, 0)
            
            self.adicionar_evento_na_interface(inicio, fim, descricao, False, excess_min)
        
        for event in sorted(eventos_locais_periodo, key=lambda x: x['start']):
            self.adicionar_evento_na_interface(
                event['start'], 
                event['end'], 
                event['description'], 
                True,
                event.get('excess_minutes', 0)  # Usar .get() para evitar KeyError
            )
        
        self.calcular_total()
        self.atualizar_tarefas_visiveis()
    
    def adicionar_evento_na_interface(self, inicio, fim, descricao, is_local, excess_minutes=0):
        """Adiciona um evento à interface, seja do calendário ou local"""
        var = tk.IntVar(value=1)
        self.check_vars.append(var)
        
        duracao = fim - inicio
        minutos = int(duracao.total_seconds() // 60)
        horas_duracao = minutos // 60
        mins_duracao = minutos % 60
        
        # Formatar duração
        if horas_duracao > 0:
            duracao_str = f"{horas_duracao}h{mins_duracao:02d}m"
        else:
            duracao_str = f"{mins_duracao}m"
        
        # Converter para horário de Brasília se for evento do calendário
        if not is_local:
            inicio_br = inicio.astimezone(self.tz_brasil)
            fim_br = fim.astimezone(self.tz_brasil)
            inicio_display = inicio_br
            fim_display = fim_br
        else:
            inicio_display = inicio
            fim_display = fim
        
        frame = ttk.Frame(self.scrollable_frame)
        frame.pack(fill='x', padx=5, pady=2)
        
        # Armazenar dados do evento no frame
        if is_local:
            frame.local_event_data = {
                'start': inicio,
                'end': fim,
                'description': descricao,
                'excess_minutes': excess_minutes
            }
        else:
            frame.event_data = (inicio, fim, descricao)
        
        # Checkbox
        cb = ttk.Checkbutton(frame, variable=var, command=self.calcular_total)
        cb.pack(side='left')

        frame.check_var = var
        
        # Ícone para identificar eventos locais
        if is_local:
            ttk.Label(frame, text="📌", font=('Arial', 10)).pack(side='left', padx=2)
        
        # Data e hora
        ttk.Label(frame, text=f"{inicio_display.strftime('%Y-%m-%d %H:%M')} → {fim_display.strftime('%H:%M')}").pack(side='left', padx=5)
        
        # Duração
        ttk.Label(frame, text=duracao_str, width=10).pack(side='left', padx=5)
        
        # Minutos excedentes
        ttk.Label(frame, text="Excedentes:").pack(side='left', padx=5)
        excess_spin = ttk.Spinbox(frame, from_=0, to=999, width=5)
        excess_spin.set(excess_minutes)
        excess_spin.pack(side='left', padx=5)
        excess_spin.bind('<FocusOut>', lambda e, f=frame: self.atualizar_minutos_excedentes(f, is_local))
        excess_spin.bind('<Return>', lambda e, f=frame: self.atualizar_minutos_excedentes(f, is_local))
        
        # Armazenar referência ao Spinbox no frame
        frame.excess_spin = excess_spin
        
        # Descrição do evento
        ttk.Label(frame, text=descricao, wraplength=400).pack(side='left', padx=5, fill='x', expand=True)
        
        # Botão para remover evento local
        if is_local:
            def remover_evento():
                if messagebox.askyesno("Confirmar", "Deseja remover este evento local permanentemente?"):
                    # Remover da lista de eventos locais
                    for i, event in enumerate(self.local_events):
                        if (event['start'] == inicio and 
                            event['end'] == fim and 
                            event['description'] == descricao):
                            del self.local_events[i]
                            self.salvar_eventos_locais()
                            break
                    # Remover da interface
                    frame.destroy()
                    self.calcular_total()
            
            ttk.Button(frame, text="×", width=2, command=remover_evento).pack(side='right', padx=5)
        
        # Armazenar minutos totais no frame
        frame.minutos_totais = minutos
    
    def atualizar_minutos_excedentes(self, frame, is_local):
        """Atualiza os minutos excedentes quando o valor é alterado"""
        try:
            excess_min = int(frame.excess_spin.get())
            
            if is_local:
                # Atualizar evento local
                event_data = frame.local_event_data
                event_data['excess_minutes'] = excess_min
                
                # Atualizar na lista principal
                for event in self.local_events:
                    if (event['start'] == event_data['start'] and 
                        event['end'] == event_data['end'] and 
                        event['description'] == event_data['description']):
                        event['excess_minutes'] = excess_min
                        break
                
                self.salvar_eventos_locais()
            else:
                # Atualizar evento do calendário
                inicio, fim, descricao = frame.event_data
                inicio_br = inicio.astimezone(self.tz_brasil)
                
                # Criar chave única para o evento
                event_key = f"{inicio_br.strftime('%Y%m%d%H%M')}_{descricao}"
                
                # Atualizar dicionário e salvar
                self.excess_minutes[event_key] = excess_min
                self.salvar_minutos_excedentes()
            
            # Recalcular totais
            self.calcular_total()
        except Exception as e:
            print(f"Erro ao atualizar minutos excedentes: {e}")
    
    def processar_calendario(self, calendario, data_inicio, data_fim):
        eventos = []
        
        # Converter para datetime no início do dia
        inicio_periodo = datetime.combine(data_inicio, datetime.min.time())
        fim_periodo = datetime.combine(data_fim, datetime.max.time())  # Fim do último dia
        
        # Converter para UTC
        inicio_periodo_utc = UTC.localize(inicio_periodo)
        fim_periodo_utc = UTC.localize(fim_periodo)
        
        for componente in calendario.walk():
            if componente.name != "VEVENT":
                continue
            
            dtstart = componente.get('dtstart').dt
            dtend = componente.get('dtend').dt
            descricao = str(componente.get('summary', 'Sem descrição')).strip()
            
            # Pular eventos cancelados
            if 'cancelado' in descricao.lower():
                continue
            
            # Converter para datetime com timezone
            if isinstance(dtstart, datetime):
                if dtstart.tzinfo is None:
                    dtstart = UTC.localize(dtstart)
            else:
                dtstart = UTC.localize(datetime.combine(dtstart, datetime.min.time()))
            
            if isinstance(dtend, datetime):
                if dtend.tzinfo is None:
                    dtend = UTC.localize(dtend)
            else:
                dtend = UTC.localize(datetime.combine(dtend, datetime.min.time()))
            
            # Processar eventos recorrentes
            if 'RRULE' in componente:
                try:
                    rrule_str = componente['RRULE'].to_ical().decode('utf-8')
                    
                    # Corrigir o UNTIL no RRULE se necessário
                    if 'UNTIL=' in rrule_str:
                        parts = rrule_str.split(';')
                        new_parts = []
                        for part in parts:
                            if part.startswith('UNTIL='):
                                until_val = part[6:]
                                if until_val.endswith('Z'):
                                    until_val = until_val[:-1]
                                try:
                                    until_dt = datetime.strptime(until_val, '%Y%m%dT%H%M%S')
                                except ValueError:
                                    until_dt = datetime.strptime(until_val, '%Y%m%d')
                                until_dt = UTC.localize(until_dt)
                                part = f"UNTIL={until_dt.strftime('%Y%m%dT%H%M%SZ')}"
                            new_parts.append(part)
                        rrule_str = ';'.join(new_parts)
                    
                    rule = rrulestr(rrule_str, dtstart=dtstart)
                    
                    # Processar EXDATEs (exceções)
                    exdates = []
                    if 'EXDATE' in componente:
                        exdate = componente['EXDATE']
                        if isinstance(exdate, list):
                            for ex in exdate:
                                exdates.extend([d.dt for d in ex.dts])
                        else:
                            exdates = [d.dt for d in exdate.dts]
                    
                    # Converter exdates para UTC
                    exdates = [UTC.localize(ex) if isinstance(ex, datetime) and ex.tzinfo is None 
                             else ex for ex in exdates]
                    
                    # Obter ocorrências dentro do período exato de 14 dias
                    for occurrence in rule.between(inicio_periodo_utc, fim_periodo_utc, inc=True):
                        if isinstance(occurrence, datetime):
                            if occurrence.tzinfo is None:
                                occurrence = UTC.localize(occurrence)
                            
                            # Verificar se não está nas exceções
                            if not any(abs((occurrence - exdate).total_seconds()) < 60 for exdate in exdates):
                                event_end = occurrence + (dtend - dtstart)
                                if event_end > occurrence:  # Verificar se a duração é válida
                                    eventos.append((occurrence, event_end, descricao))
                    
                except Exception as e:
                    print(f"Erro ao processar evento recorrente {descricao}: {str(e)}")
                    continue
            else:
                # Evento único - verificar se está dentro do período de 14 dias
                if inicio_periodo_utc <= dtstart <= fim_periodo_utc:
                    eventos.append((dtstart, dtend, descricao))
        
        return eventos
    
    def aplicar_filtros(self, eventos):
        # Primeiro, remove eventos cancelados
        eventos_validos = [e for e in eventos if 'cancelado' not in e[2].lower()]
        
        # Agrupa eventos por data
        eventos_por_data = {}
        for evento in eventos_validos:
            data = evento[0].astimezone(self.tz_brasil).date()
            eventos_por_data.setdefault(data, []).append(evento)
        
        # Processa cada dia para resolver conflitos
        eventos_filtrados = []
        
        for data, eventos_dia in eventos_por_data.items():
            # Ordena eventos por horário de início
            eventos_dia.sort(key=lambda x: x[0])
            
            i = 0
            while i < len(eventos_dia):
                current_start, current_end, current_desc = eventos_dia[i]
                max_duration = current_end - current_start
                final_desc = current_desc
                j = i + 1
                
                # Encontra todos os eventos sobrepostos
                while j < len(eventos_dia):
                    next_start, next_end, next_desc = eventos_dia[j]
                    
                    # Verifica se há sobreposição
                    if next_start >= current_end:
                        break
                    
                    # Atualiza para pegar o término mais tarde
                    current_end = max(current_end, next_end)
                    
                    # Verifica qual tem maior duração para pegar a descrição
                    next_duration = next_end - next_start
                    if next_duration > max_duration:
                        max_duration = next_duration
                        final_desc = next_desc
                    
                    j += 1
                
                # Adiciona o evento consolidado
                eventos_filtrados.append((current_start, current_end, final_desc))
                i = j
        
        return eventos_filtrados
    
    def calcular_total(self):
        total_minutos = 0
        total_excess = 0
        eventos_selecionados = 0
        
        for frame in self.scrollable_frame.winfo_children():
            if hasattr(frame, 'minutos_totais'):
                for widget in frame.winfo_children():
                    if isinstance(widget, ttk.Checkbutton):
                        if widget.instate(['selected']):
                            total_minutos += frame.minutos_totais
                            try:
                                total_excess += int(frame.excess_spin.get())
                            except:
                                pass
                            eventos_selecionados += 1
                        break
        
        # Formatar igual à lista (ex: 2h30m)
        horas = total_minutos // 60
        minutos = total_minutos % 60
        
        if horas > 0:
            total_str = f"{horas}h{minutos:02d}m"
        else:
            total_str = f"{minutos}m"
        
        # Formatar minutos excedentes
        excess_horas = total_excess // 60
        excess_mins = total_excess % 60
        
        if excess_horas > 0:
            excess_str = f"{excess_horas}h{excess_mins:02d}m"
        else:
            excess_str = f"{excess_mins}m"
        
        # Tempo previsto
        horas = total_minutos // 60
        minutos = total_minutos % 60
        tempo_previsto_str = f"{horas}h{minutos:02d}m" if horas > 0 else f"{minutos}m"

        # Excedidos
        excess_horas = total_excess // 60
        excess_mins = total_excess % 60
        excedidos_str = f"{excess_horas}h{excess_mins:02d}m" if excess_horas > 0 else f"{excess_mins}m"

        # Total final
        total_final = total_minutos + total_excess
        final_horas = total_final // 60
        final_mins = total_final % 60
        tempo_total_str = f"{final_horas}h{final_mins:02d}m" if final_horas > 0 else f"{final_mins}m"

        self.resultado_label.config(
            text=f"Tempo previsto: {tempo_previsto_str} | Excedidos: {excedidos_str} | Tempo Total: {tempo_total_str}"
        )


if __name__ == "__main__":
    root = tk.Tk()
    app = CalendarTrackerApp(root)
    root.mainloop()
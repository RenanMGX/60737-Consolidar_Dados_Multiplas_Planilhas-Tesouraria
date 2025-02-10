import os
import json
from dependencies.functions import P
from datetime import datetime


class LogInformativo:
    @property
    def file_path(self):
        return os.path.join(os.getcwd(), 'informativoLog.json')
    
    def __init__(self):
        if not os.path.exists(self.file_path):
            with open(self.file_path, 'w', encoding='utf-8') as f:
                json.dump([], f)
                
    def get(self) -> list:
        with open(self.file_path, 'r', encoding='utf-8') as f:
            try:
                return json.load(f)
            except:
                return []
        
    def add(self, message:str) -> None:
        logs = self.get()
        date_tag = datetime.now().strftime('[%Y-%m-%d %H:%M:%S]')
        logs.append(f"{date_tag} - {message}")
        
        with open(self.file_path, 'w', encoding='utf-8') as f:
            json.dump(logs, f)
            
    def clear(self) -> None:
        with open(self.file_path, 'w', encoding='utf-8') as f:
            json.dump([], f)
            
if __name__ == "__main__":
    log = LogInformativo()
    log.add("Teste")
    print(log.get())
    #log.clear()
    print(log.get())
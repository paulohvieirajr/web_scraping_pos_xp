from typing import List

class Artigo():
    titulo:str = None
    paragrafos:List[str] = []

    def __init__(self):
        self.titulo = None
        self.paragrafos = []
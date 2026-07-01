"""
Модуль семантического поиска: сравнение текста с коллекцией документов.
Использует sentence-transformers для получения эмбеддингов и косинусное сходство.
"""

import torch
import numpy as np
from transformers import AutoTokenizer, AutoModel

class SemanticSearch:
    def __init__(self, model_name="sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2"):
        self.tok = AutoTokenizer.from_pretrained(model_name)
        self.model = AutoModel.from_pretrained(model_name)
        self.docs = []
        self.emb = None

    def _encode(self, texts):
        """Получение эмбеддингов (mean pooling)."""
        toks = self.tok(texts, padding=True, truncation=True, return_tensors="pt", max_length=512)
        with torch.no_grad():
            out = self.model(**toks)
        mask = toks["attention_mask"].unsqueeze(-1).expand(out.last_hidden_state.size()).float()
        return (out.last_hidden_state * mask).sum(1) / mask.sum(1).clamp(min=1e-9)

    def fit(self, documents):
        """Индексация коллекции документов."""
        self.docs = documents
        self.emb = self._encode(documents)
        self.emb = self.emb / self.emb.norm(dim=1, keepdim=True)

    def search(self, query, top_k=1):
        """Поиск наиболее похожих документов. Возвращает [(score, doc)]."""
        q = self._encode([query])
        q = q / q.norm(dim=1, keepdim=True)
        scores = (q @ self.emb.T).squeeze().numpy()
        if scores.ndim == 0:
            scores = np.array([scores])
        idx = np.argsort(scores)[::-1][:top_k]
        return [(float(scores[i]), self.docs[i]) for i in idx]


# === Пример использования ===
if __name__ == "__main__":
    docs = [
        "Python — язык программирования общего назначения",
        "JavaScript используется для веб-разработки",
        "SQL — язык запросов к базам данных",
        "PyTorch — библиотека для машинного обучения",
    ]

    ss = SemanticSearch()
    ss.fit(docs)

    queries = [
        "какой язык для веба",
        "нейросети и машинное обучение",
        "базы данных и запросы",
    ]

    for q in queries:
        results = ss.search(q, top_k=2)
        print(f"\nЗапрос: '{q}'")
        for score, doc in results:
            print(f"  {score:.3f} | {doc}")

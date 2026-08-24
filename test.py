import requests

r = requests.post(
    "http://103.130.147.241:44609/v1/rerank",
    json={
        "model": "BAAI/bge-reranker-v2-m3",
        "query": "test query",
        "documents": ["a relevant test document", "unrelated text"],
    },
    timeout=15,
)
print(r.status_code)
print(r.json())
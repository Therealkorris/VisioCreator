import logging
import json
from langchain_ollama import ChatOllama
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain_community.document_loaders import WebBaseLoader
from langchain_community.vectorstores import SKLearnVectorStore
from langchain_nomic.embeddings import NomicEmbeddings
from langchain.schema import Document
from typing import List

logging.basicConfig(level=logging.INFO)

# Initialize LLM
local_llm = "llama3.2:3b-instruct-fp16"
llm = ChatOllama(model=local_llm, temperature=0)
llm_json_mode = ChatOllama(model=local_llm, temperature=0, format="json")

# Load Documents from URLs
urls = [
    "https://lilianweng.github.io/posts/2023-06-23-agent/",
    "https://lilianweng.github.io/posts/2023-03-15-prompt-engineering/",
    "https://lilianweng.github.io/posts/2023-10-25-adv-attack-llm/",
]

retriever = None
try:
    loaded_docs = []
    for url in urls:
        loader = WebBaseLoader(url)
        docs = loader.load()
        for doc in docs:
            if isinstance(doc, Document):
                loaded_docs.append(doc)
    # Split documents for VectorDB
    text_splitter = RecursiveCharacterTextSplitter.from_tiktoken_encoder(chunk_size=1000, chunk_overlap=200)
    doc_splits = text_splitter.split_documents(loaded_docs)
    # Add to vectorDB
    vectorstore = SKLearnVectorStore.from_documents(
        documents=doc_splits,
        embedding=NomicEmbeddings(model="nomic-embed-text-v1.5", inference_mode="local"),
    )
    retriever = vectorstore.as_retriever(k=3)
except Exception as e:
    logging.error(f"Failed to load documents from URLs: {e}")

# Router Prompt
router_instructions = """
You are an expert at routing a user question to a vectorstore or web search.

The vectorstore contains documents related to agents, prompt engineering, and adversarial attacks.

Use the vectorstore for questions on these topics. For all else, and especially for current events, use web-search.

Return JSON with single key, datasource, that is 'websearch' or 'vectorstore' depending on the question.
"""

# Test Router
def test_router():
    questions = [
        "Who is favored to win the NFC Championship game in the 2024 season?",
        "What are the models released today for llama3.2?",
        "What are the types of agent memory?"
    ]
    for question in questions:
        response = llm_json_mode.invoke([
            {"role": "system", "content": router_instructions},
            {"role": "user", "content": question}
        ])
        print(json.loads(response.content))

# Test Retrieval Grader

doc_grader_instructions = """
You are a grader assessing relevance of a retrieved document to a user question.

If the document contains keyword(s) or semantic meaning related to the question, grade it as relevant.
"""

doc_grader_prompt = """
Here is the retrieved document: 

 {document} 

 Here is the user question: 

 {question}. 

This carefully and objectively assess whether the document contains at least some information that is relevant to the question.

Return JSON with single key, binary_score, that is 'yes' or 'no' score to indicate whether the document contains at least some information that is relevant to the question.
"""

# Test Retrieval Grader
def test_retrieval_grader():
    if retriever is None:
        print("Retriever is not initialized. Skipping test_retrieval_grader.")
        return
    question = "What is Chain of thought prompting?"
    docs = retriever.invoke(question)
    if not docs:
        print("No documents retrieved. Skipping test_retrieval_grader.")
        return
    doc_txt = docs[0].page_content
    doc_grader_prompt_formatted = doc_grader_prompt.format(document=doc_txt, question=question)
    response = llm_json_mode.invoke([
        {"role": "system", "content": doc_grader_instructions},
        {"role": "user", "content": doc_grader_prompt_formatted}
    ])
    print(json.loads(response.content))

# Test Generation
rag_prompt = """
You are an assistant for question-answering tasks. 

Here is the context to use to answer the question:

{context} 

Think carefully about the above context. 

Now, review the user question:

{question}

Provide an answer to this question using only the above context. 

Use three sentences maximum and keep the answer concise.

Answer:
"""

def format_docs(docs: List[Document]) -> str:
    return "\n\n".join(doc.page_content for doc in docs)

# Test Generation
def test_generation():
    if retriever is None:
        print("Retriever is not initialized. Skipping test_generation.")
        return
    question = "What is Chain of thought prompting?"
    docs = retriever.invoke(question)
    if not docs:
        print("No documents retrieved. Skipping test_generation.")
        return
    docs_txt = format_docs(docs)
    rag_prompt_formatted = rag_prompt.format(context=docs_txt, question=question)
    response = llm.invoke([{ "role": "user", "content": rag_prompt_formatted }])
    print(response.content)

if __name__ == "__main__":
    test_router()
    test_retrieval_grader()
    test_generation()
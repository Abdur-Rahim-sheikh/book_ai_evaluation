## How to run

### Pre-requisites
- Download ollama
  - either directly to your computer from the [releases page](https://ollama.com/download)
  - Or get the docker image from [Docker Hub](https://hub.docker.com/r/ollama/ollama)
    - `docker run -d -v ollama:/root/.ollama -p 11434:11434 --name ollama ollama/ollama`
- Install poetry
  - from [official page](https://python-poetry.org/docs/#installing-with-the-official-installer)
    - i.e `curl -sSL https://install.python-poetry.org | python3 -`


### Running the project

- Clone the repository `git clone https://github.com/Abdur-Rahim-sheikh/book_ai_evaluation.git`
- `cd book_ai_evaluation`
- Set your environment variables, if you want to use chatGPT and change the main.py files import option
- `python main.py`
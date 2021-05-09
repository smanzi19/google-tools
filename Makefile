init:
	pyenv install -s 3.8.5
	pyenv virtualenv 3.8.5 google-tools
install:
	poetry install
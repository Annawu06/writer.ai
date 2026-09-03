all:
	./build.sh
	/opt/libreoffice26.2/program/unopkg add -f writer.ai.oxt

test:
	/opt/libreoffice26.2/program/python -m unittest discover -s tests -p 'test_*.py'

PYTHON ?= $(if $(wildcard /opt/libreoffice26.2/program/python),/opt/libreoffice26.2/program/python,/usr/bin/python3)
UNOPKG ?= $(if $(wildcard /opt/libreoffice26.2/program/unopkg),/opt/libreoffice26.2/program/unopkg,unopkg)

all:
	./build.sh
	$(UNOPKG) add -f writer.ai.oxt

test:
	$(PYTHON) -m unittest discover -s tests -p 'test_*.py'

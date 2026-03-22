PYTHON ?= python3
MPLCONFIGDIR ?= /tmp/matplotlib

.PHONY: deps test tests compile clean

deps:
	$(PYTHON) -m pip install -r requirements.txt

test tests:
	MPLCONFIGDIR=$(MPLCONFIGDIR) $(PYTHON) -m unittest discover -s tests -v

compile:
	$(PYTHON) -m compileall ./

clean:
	find . -type d -name "__pycache__" -prune -exec rm -rf {} +

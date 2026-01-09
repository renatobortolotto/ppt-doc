CONTAINER=genaigke-mt-base-extratos-dados-deb
MOCKED_APPLICATION_NAME=$CONTAINER
PYTHONCMD=python3
PYTHON36_FOUND=$(shell command -v python36 2> /dev/null)

AMBIENTE ?= prod

ifeq ($(AMBIENTE), prod)
	PYTEST_UNIT_EXCLUDE :=
	PIP_INSTALL_FLAGS := --user
endif

ifeq ($(AMBIENTE), terceiros)
	PYTEST_UNIT_EXCLUDE := -m "not hbase"
	PIP_INSTALL_FLAGS :=
endif

ifndef PYTHON36_FOUND
PYTHONCMD=python36
endif

compile: deps
	$(PYTHONCMD) -m compileall ./

lint: devdeps
	flake8 . || true 

tests:
	exit 0 

unit-tests: devdeps delete-coverage-report
	$(PYTHONCMD) -m pytest tests/unit --cov=./ --cov-report=term-missing:skip-covered --cov-report=xml --cov-fail-under=80 $(PYTEST_UNIT_EXCLUDE)

delete-coverage-report:
	rm -f coverage.xml

integration-tests:
	exit 0

deps: 
	$(PYTHONCMD) -m pip install $(PIP_INSTALL_FLAGS) -r requirements.txt

devdeps:
	$(PYTHONCMD) -m pip install $(PIP_INSTALL_FLAGS) -r requirements-dev.txt

build-wheel: devdeps
	$(PYTHONCMD) setup.py sdist bdist_wheel

clean:
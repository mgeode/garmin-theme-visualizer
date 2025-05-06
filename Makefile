APP_NAME:=garmin-theme-visualizer

.pony: help
help: 
	@echo ""
	@echo "+++ Makefile help +++"
	@grep -F -h "###" $(MAKEFILE_LIST) | grep -F -v grep | sed -e 's/\\$$//' | sed -e 's/:.*##/:\t\t/'
	@echo ""
	@echo "+++ App-Settings +++"
	@python3 ./${APP_NAME}.py -h
	@echo ""
  

install: ### install dependencies
	@python3 -m pip install -r requirements.txt

run: ### run application
	@echo "Starting app {APP_NAME}"
	@python3 ./${APP_NAME}.py -i themes/mgeo/mgeo-v206.kmtf
set GO111MODULE = on
set GOOS = windows
go build -o build/exporter-win.exe ./src/
set GOOS = linux
go build -o build/exporter-linux ./src/
set GOOS = darwin
go build -o build/exporter-darwin ./src/
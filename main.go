package main

import (
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"strings"

	"github.com/atotto/clipboard"
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

func main() {
	installerPath := "C:\\Windows\\Installer"
	authorName := "SolarWinds"

	err := filepath.Walk(installerPath, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}
		if !info.IsDir() && strings.HasSuffix(info.Name(), ".msi") {
			properties, err := getMsiProperties(path)
			if err != nil {
				fmt.Printf("Error processing file: %s\n", path)
				return nil
			}
			if properties["Manufacturer"] == authorName {
				fmt.Printf("File: %s\n", path)
				uninstallMsi(path)
				clipboard.WriteAll("burn/purify")
				runExecutable("C:\\SolarWinds.Licensing.Reset.exe")
				runExecutable("C:\\SolarWinds-Dameware-MRC-64bit.exe")
				fmt.Println("Done!")
			}
		}
		return nil
	})

	if err != nil {
		fmt.Println("Error walking the path:", err)
	}
}

func getMsiProperties(path string) (map[string]string, error) {
	properties := make(map[string]string)
	ole.CoInitialize(0)
	defer ole.CoUninitialize()

	unknown, err := oleutil.CreateObject("WindowsInstaller.Installer")
	if err != nil {
		return properties, err
	}
	defer unknown.Release()

	msi, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return properties, err
	}
	defer msi.Release()

	db, err := oleutil.CallMethod(msi, "OpenDatabase", path, 0)
	if err != nil {
		return properties, err
	}
	defer db.Clear()

	view, err := oleutil.CallMethod(db.ToIDispatch(), "OpenView", "SELECT Property, Value FROM Property")
	if err != nil {
		return properties, err
	}
	defer view.Clear()

	oleutil.MustCallMethod(view.ToIDispatch(), "Execute")

	for {
		record, err := oleutil.CallMethod(view.ToIDispatch(), "Fetch")
		if err != nil || record.Val == 0 {
			break
		}
		propertyName := oleutil.MustGetProperty(record.ToIDispatch(), "StringData", 1).ToString()
		propertyValue := oleutil.MustGetProperty(record.ToIDispatch(), "StringData", 2).ToString()
		properties[propertyName] = propertyValue
		record.Clear()
	}

	return properties, nil
}

func uninstallMsi(path string) {
	cmd := exec.Command("msiexec.exe", "/x", path)
	err := cmd.Run()
	if err != nil {
		fmt.Printf("Error uninstalling MSI: %s\n", err)
	}
}

func runExecutable(path string) {
	cmd := exec.Command(path)
	err := cmd.Start()
	if err != nil {
		fmt.Printf("Error starting executable: %s\n", err)
		return
	}
	err = cmd.Wait()
	if err != nil {
		fmt.Printf("Error running executable: %s\n", err)
	}
}

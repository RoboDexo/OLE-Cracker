package main

import (
    "fmt"
    "io/ioutil"
    "log"
    "path/filepath"
    "os"
    "strings"
    "time"

    "github.com/360EntSecGroup-Skylar/excelize/v2"
    "github.com/360EntSecGroup-Skylar/go-getting-started-with-go/go-docx"
    "github.com/360EntSecGroup-Skylar/go-getting-started-with-go/go-pptx"
    "github.com/360EntSecGroup-Skylar/go-getting-started-with-go/go-xls"
    "github.com/minio/minio"
    "github.com/minio/minio-go/v7/pkg/credentials"
    "github.com/minio/minio-go/v7/pkg/errors"
)

func main() {
    // Set the path to the Office document
    filePath := "path_to_your_office_file.docx"

    // Extract the hash from the Office document
    hash, err := extractHash(filePath)
    if err != nil {
        log.Fatal(err)
    }

    // Brute-force the hash
    if len(os.Args) > 1 && os.Args[1] == "-d" {
        // Use a dictionary for brute-forcing
        dictPath := "path_to_your_wordlist.txt"
        bruteForceHash(hash, dictPath)
    } else {
        // Perform a raw brute-force of all characters
        bruteForceHash(hash, "")
    }
}

func extractHash(filePath string) (string, error) {
    // Open the file
    file, err := ioutil.ReadFile(filePath)
    if err != nil {
        return "", err
    }

    // Determine the file type
    fileExt := filepath.Ext(filePath)
    switch fileExt {
    case ".docx":
        return extractDocxHash(file)
    case ".doc":
        return extractDocHash(file)
    case ".xls", ".xlsx":
        return extractXlsHash(file)
    case ".ppt", ".pptx":
        return extractPptHash(file)
    default:
        return "", fmt.Errorf("unsupported file type: %s", fileExt)
    }
}

func extractDocxHash(file []byte) (string, error) {
    // Use the go-docx library to extract the hash
    doc, err := go_docx.NewDocx(file)
    if err != nil {
        return "", err
    }

    // Get the document's hash
    hash, err := doc.GetHash()
    if err != nil {
        return "", err
    }

    return hash, nil
}

func extractDocHash(file []byte) (string, error) {
    // TO DO: implement doc hash extraction
    return "", nil
}

func extractXlsHash(file []byte) (string, error) {
    // Use the go-xls library to extract the hash
    xls, err := xls.NewXLS(file)
    if err != nil {
        return "", err
    }

    // Get the workbook's hash
    hash, err := xls.GetHash()
    if err != nil {
        return "", err
    }

    return hash, nil
}

func extractPptHash(file []byte) (string, error) {
    // Use the go-pptx library to extract the hash
    ppt, err := pptx.NewPPTX(file)
    if err != nil {
        return "", err
    }

    // Get the presentation's hash
    hash, err := ppt.GetHash()
    if err != nil {
        return "", err
    }

    return hash, nil
}

func bruteForceHash(hash string, dictPath string) {
    if dictPath != "" {
        // Use a dictionary for brute-forcing
        file, err := ioutil.ReadFile(dictPath)
        if err != nil {
            log.Fatal(err)
        }

        words := strings.Split(string(file), "\n")
        for _, word := range words {
            if hash == hashFunc(word) {
                log.Println("Found password:", word)
                return
            }
        }
    } else {
        // Perform a raw brute-force of all characters
        for i := 0; i < 256; i++ {
            for j := 0; j < 256; j++ {
                for k := 0; k < 256; k++ {
                    password := fmt.Sprintf("%c%c%c", byte(i), byte(j), byte(k))
                    if hash == hashFunc(password) {
                        log.Println("Found password:", password)
                        return
                    }
                }
            }
        }
    }
}

func hashFunc(password string) string {
    // TO DO: implement your hash// TO DO: implement your hash function
    // For example, you can use the MD5 hash function
    md5 := md5.New()
    md5.Write([]byte(password))
    hash := md5.Sum(nil)
    return fmt.Sprintf("%x", hash)
}

func main() {
    // Set the path to the Office document
    filePath := "path_to_your_office_file.docx"

    // Extract the hash from the Office document
    hash, err := extractHash(filePath)
    if err != nil {
        log.Fatal(err)
    }

    // Brute-force the hash
    if len(os.Args) > 1 && os.Args[1] == "-d" {
        // Use a dictionary for brute-forcing
        dictPath := "path_to_your_wordlist.txt"
        bruteForceHash(hash, dictPath)
    } else {
        // Perform a raw brute-force of all characters
        bruteForceHash(hash, "")
    }
}

func main() {
    // Set the path to the Office document
    filePath := "path_to_your_office_file.docx"

    // Extract the hash from the Office document
    hash, err := extractHash(filePath)
    if err != nil {
        log.Fatal(err)
    }

    // Brute-force the hash
    if len(os.Args) > 1 && os.Args[1] == "-d" {
        // Use a dictionary for brute-forcing
        dictPath := "path_to_your_wordlist.txt"
        bruteForceHash(hash, dictPath)
    } else {
        // Perform a raw brute-force of all characters
        bruteForceHash(hash, "")
    }
}

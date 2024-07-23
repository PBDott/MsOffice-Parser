#include <stdio.h>
#include <string.h>
#include <stdint.h>

// 함수 선언
int OldOfficeFile(const char* filePath);
int NewOfficeFile(const char* filePath);

int OldOfficeFile(const char* filePath) {
    FILE* file = fopen(filePath, "rb");

    if (file == NULL) {
        perror("파일을 열 수 없습니다.");
        return 0; // 오류 코드 반환
    }

    char signature[8];
    fread(signature, 1, 8, file);

    unsigned char buffer[7]; // 7바이트 패턴    

    // 구 버전 Office 파일 시그니처 확인
    if (memcmp(signature, "\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1", 8) == 0) {
        printf("구버전 MS Office, CFBF 형식입니다. (PPT, DOC, XLS)\n");

        //doc 스트림 주소
        fseek(file, 0x5F80, SEEK_SET);

        //doc 파일 내용을 읽어옴
        fread(buffer, sizeof(unsigned char), 7, file);

        //doc 바이너리 패턴 비교
        if (buffer[0] == 0x57 && buffer[2] == 0x6F && buffer[4] == 0x72 && buffer[6] == 0x64) {
            printf("이 파일은 'doc' 파일입니다.\n");
        } 
        
        //ppt 스트림 주소
        fseek(file, 0x480, SEEK_SET);

        //ppt 파일 내용을 읽어옴
        fread(buffer, sizeof(unsigned char), 7, file);

        if(buffer[0] == 0x50 && buffer[2] == 0x6F && buffer[4] == 0x77 && buffer[6] == 0x65){
            printf("이 파일은 'ppt' 파일입니다.\n");
        }

        //xls 스트림 주소
        fseek(file, 0x480, SEEK_SET);

        //xls 파일 내용을 읽어옴
        fread(buffer, sizeof(unsigned char), 7, file);
   
        if(buffer[0] == 0x57 && buffer[2] == 0x6F && buffer[4] == 0x72 && buffer[6] == 0x6B){
            printf("이 파일은 'xls' 파일입니다.\n");
        }


        return 1; // 구 버전 Office 파일
    }

    fclose(file);
    return 0; // 구 버전 Office 파일이 아님
}

int NewOfficeFile(const char* filePath) {
    FILE* file = fopen(filePath, "rb");
    
    if (file == NULL) {
        perror("파일을 열 수 없습니다.");
        return 0; // 오류 코드 반환
    }

    char signature[8];
    fread(signature, 1, 8, file);

    unsigned char buffer[16]; // 7바이트 패턴

    // 신 버전 Office 파일 시그니처 확인
    if (memcmp(signature, "\x50\x4B\x03\x04\x14\x00\x06\x00", 8) == 0) {
        printf("신버전 MS Office, OOXML 형식입니다. (PPTX, DOCX, XLSX)\n");

        //docx 스트림 주소
        fseek(file, 0x6D0, SEEK_SET);

        //docx 파일 내용을 읽어옴
        fread(buffer, sizeof(unsigned char), 16, file);

        //docx 바이너리 패턴 비교
        if (buffer[1] == 0x77 && buffer[2] == 0x6F && buffer[3] == 0x72 && buffer[4] == 0x64) {
            printf("이 파일은 'docx' 파일입니다.\n");
        } 

        //pptx 스트림 주소
        fseek(file, 0x990, SEEK_SET);

        //pptx 파일 내용을 읽어옴
        fread(buffer, sizeof(unsigned char), 16, file);

        if(buffer[2] == 0x70 && buffer[3] == 0x70 && buffer[4] == 0x74){
            printf("이 파일은 'pptx' 파일입니다.\n");
        }

        //xlsx 스트림 주소
        fseek(file, 0x6D0, SEEK_SET);

        //xlsx 파일 내용을 읽어옴
        fread(buffer, sizeof(unsigned char), 16, file);
   
        if(buffer[10] == 0x78 && buffer[11]){
            printf("이 파일은 'xlsx' 파일입니다.\n");
        }

        return 1; // 신 버전 Office 파일
    }

    fclose(file);
    return 0; // 신 버전 Office 파일이 아님
}

int main() {
    const char* filePath = "C:\\Users\\admin\\Desktop\\test\\test.xlsx"; // 해당 MS Office 경로
    
    printf("\n\n<---------------- 파일 정보 ---------------->\n\n");

    if (OldOfficeFile(filePath)) {
        printf("정상적으로 출력되었습니다.\n\n");
    } else if (NewOfficeFile(filePath)) {
        printf("정상적으로 출력되었습니다.\n\n");
    } else {
        printf("알 수 없는 파일 형식입니다.\n");
    }

    return 0;
}


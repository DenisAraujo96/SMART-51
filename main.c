#define UNICODE
#define _UNICODE
#include <windows.h>
#include <winscard.h>
#include <iostream>
#include <fstream>
#include <sstream>
#include <vector>
#include <string>
#include <iomanip>
#include <algorithm>
#include <cmath>

#pragma comment(lib, "winscard.lib")

// ---------------- CONFIGURAÇÃO ----------------
const wchar_t* PRINTER_NAME = nullptr;     // nullptr = impressora padrão
const char* INPUT_CSV = "input.csv";
const char* OUTPUT_CSV = "output.csv";

const double CARD_W_MM = 86.0;
const double CARD_H_MM = 54.0;

struct Field { double x_mm; double y_mm; int font_pt; };
Field FIELD_MATRICULA = { 10.0, 10.0, 18 };
Field FIELD_UID = { 10.0, 30.0, 12 };

const std::vector<std::wstring> PREFERRED_READER_TERMS = {
    L"duali", L"idp", L"smart", L"omnikey", L"dual" L"usbccid" L"wudf"
};
// -----------------------------------------------

static int mm_to_pixels(HDC hdc, double mm, bool horizontal) {
    int dpi = horizontal ? GetDeviceCaps(hdc, LOGPIXELSX) : GetDeviceCaps(hdc, LOGPIXELSY);
    double inches = mm / 25.4;
    return (int)std::round(inches * dpi);
}

static std::wstring to_wstring_acp(const std::string& s) {
    if (s.empty()) return std::wstring();
    int n = MultiByteToWideChar(CP_ACP, 0, s.c_str(), -1, NULL, 0);
    if (n <= 0) return std::wstring();
    std::wstring r(n, L'\0');
    MultiByteToWideChar(CP_ACP, 0, s.c_str(), -1, &r[0], n);
    if (!r.empty() && r.back() == L'\0') r.pop_back();
    return r;
}

static std::string remove_spaces_w_to_ascii(const std::wstring& ws) {
    std::string out;
    for (wchar_t c : ws) {
        if (c == L' ') continue;
        out.push_back(static_cast<char>(c));
    }
    return out;
}

static std::wstring find_preferred_reader(const std::vector<std::wstring>& readers) {
    for (const auto& term : PREFERRED_READER_TERMS) {
        std::wstring term_l = term;
        std::transform(term_l.begin(), term_l.end(), term_l.begin(), ::towlower);
        for (const auto& r : readers) {
            std::wstring r_l = r;
            std::transform(r_l.begin(), r_l.end(), r_l.begin(), ::towlower);
            if (r_l.find(term_l) != std::wstring::npos) return r;
        }
    }
    return std::wstring();
}

static bool read_uid_pcsc_auto_wait(std::wstring& uidHexOut, std::wstring& readerUsed, std::string& errMsg, DWORD waitTimeoutMs = 10000) {
    SCARDCONTEXT hContext = 0;
    LONG rv = SCardEstablishContext(SCARD_SCOPE_USER, NULL, NULL, &hContext);
    if (rv != SCARD_S_SUCCESS) { errMsg = "SCardEstablishContext falhou"; return false; }

    DWORD readersLen = 0;
    rv = SCardListReadersW(hContext, NULL, NULL, &readersLen);
    if (rv != SCARD_S_SUCCESS || readersLen == 0) { errMsg = "Nenhum leitor encontrado"; SCardReleaseContext(hContext); return false; }

    std::vector<wchar_t> buf(readersLen);
    rv = SCardListReadersW(hContext, NULL, buf.data(), &readersLen);
    if (rv != SCARD_S_SUCCESS) { errMsg = "SCardListReadersW falhou"; SCardReleaseContext(hContext); return false; }

    std::vector<std::wstring> readers;
    wchar_t* p = buf.data();
    while (p && *p) { readers.emplace_back(p); p += wcslen(p) + 1; }
    if (readers.empty()) { errMsg = "Nenhum leitor extraído"; SCardReleaseContext(hContext); return false; }

    std::wstring chosen = find_preferred_reader(readers);
    if (chosen.empty()) chosen = readers[0];
    readerUsed = chosen;

    SCARD_READERSTATEW readerState;
    ZeroMemory(&readerState, sizeof(readerState));
    readerState.szReader = chosen.c_str();
    readerState.dwCurrentState = SCARD_STATE_UNAWARE;

    rv = SCardGetStatusChangeW(hContext, waitTimeoutMs, &readerState, 1);
    if (rv != SCARD_S_SUCCESS) { errMsg = "SCardGetStatusChangeW falhou"; SCardReleaseContext(hContext); return false; }

    if (!(readerState.dwEventState & SCARD_STATE_PRESENT)) { errMsg = "Nenhum cartao presente"; SCardReleaseContext(hContext); return false; }

    SCARDHANDLE hCard;
    DWORD dwActiveProtocol = 0;
    rv = SCardConnectW(hContext, chosen.c_str(), SCARD_SHARE_SHARED, SCARD_PROTOCOL_T0 | SCARD_PROTOCOL_T1, &hCard, &dwActiveProtocol);
    if (rv != SCARD_S_SUCCESS) { errMsg = "SCardConnectW falhou"; SCardReleaseContext(hContext); return false; }

    BYTE getUidApdu[] = { 0xFF,0xCA,0x00,0x00,0x00 };
    BYTE recvBuf[258]; DWORD recvLen = sizeof(recvBuf);
    const SCARD_IO_REQUEST* pioSendPci = (dwActiveProtocol == SCARD_PROTOCOL_T0) ? SCARD_PCI_T0 : SCARD_PCI_T1;

    rv = SCardTransmit(hCard, pioSendPci, getUidApdu, sizeof(getUidApdu), NULL, recvBuf, &recvLen);
    if (rv != SCARD_S_SUCCESS || recvLen < 2) { errMsg = "SCardTransmit falhou ou resposta curta"; SCardDisconnect(hCard, SCARD_LEAVE_CARD); SCardReleaseContext(hContext); return false; }

    size_t dataLen = recvLen - 2;
    std::wostringstream ws;
    ws << std::uppercase << std::hex << std::setfill(L'0');
    for (size_t i = 0; i < dataLen; ++i) { if (i) ws << L' '; ws << std::setw(2) << (int)recvBuf[i]; }
    uidHexOut = ws.str();

    SCardDisconnect(hCard, SCARD_LEAVE_CARD);
    SCardReleaseContext(hContext);
    return true;
}

static bool print_card_gdi(const std::wstring& matricula, const std::wstring& uidHex) {
    DOCINFOW docInfo; ZeroMemory(&docInfo, sizeof(docInfo)); docInfo.cbSize = sizeof(docInfo); docInfo.lpszDocName = L"CardPrintJob";
    HDC hdc = CreateDCW(L"WINSPOOL", PRINTER_NAME, NULL, NULL);
    if (!hdc) return false;
    if (StartDocW(hdc, &docInfo) <= 0) { DeleteDC(hdc); return false; }
    if (StartPage(hdc) <= 0) { EndDoc(hdc); DeleteDC(hdc); return false; }

    int card_w_px = mm_to_pixels(hdc, CARD_W_MM, true);
    int card_h_px = mm_to_pixels(hdc, CARD_H_MM, false);
    RECT rc = { 0,0,card_w_px,card_h_px };
    FrameRect(hdc, &rc, (HBRUSH)GetStockObject(BLACK_BRUSH));

    int dpiY = GetDeviceCaps(hdc, LOGPIXELSY);
    HFONT hFontMat = CreateFontW(-MulDiv(FIELD_MATRICULA.font_pt, dpiY, 72), 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Arial");
    HFONT hFontUid = CreateFontW(-MulDiv(FIELD_UID.font_pt, dpiY, 72), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Arial");

    int mx = mm_to_pixels(hdc, FIELD_MATRICULA.x_mm, true);
    int my = mm_to_pixels(hdc, FIELD_MATRICULA.y_mm, false);
    int ux = mm_to_pixels(hdc, FIELD_UID.x_mm, true);
    int uy = mm_to_pixels(hdc, FIELD_UID.y_mm, false);

    HFONT old = (HFONT)SelectObject(hdc, hFontMat);
    SetTextColor(hdc, RGB(0, 0, 0));
    SetBkMode(hdc, TRANSPARENT);
    TextOutW(hdc, mx, my, matricula.c_str(), (int)matricula.size());

    SelectObject(hdc, hFontUid);
    TextOutW(hdc, ux, uy, uidHex.c_str(), (int)uidHex.size());

    SelectObject(hdc, old);
    DeleteObject(hFontMat);
    DeleteObject(hFontUid);

    EndPage(hdc);
    EndDoc(hdc);
    DeleteDC(hdc);
    return true;
}

int wmain(int argc, wchar_t* argv[]) {
    std::wcout << L"=== SmartCard UID + Print ===\n";
    std::ifstream in(INPUT_CSV);
    if (!in.is_open()) { std::wcerr << L"Erro ao abrir input.csv\n"; return 1; }
    std::vector<std::string> matriculas; std::string line;
    while (std::getline(in, line)) { if (!line.empty()) matriculas.push_back(line); } in.close();
    if (matriculas.empty()) { std::wcerr << L"Nenhuma matricula\n"; return 1; }

    std::ofstream out(OUTPUT_CSV, std::ios::app);
    if (!out.is_open()) { std::wcerr << L"Erro ao abrir output.csv\n"; return 1; }

    for (size_t i = 0; i < matriculas.size(); ++i) {
        const std::string mat = matriculas[i];
        std::wcout << L"Processando matricula: " << to_wstring_acp(mat) << L"\nPressione Enter com o cartAo no leitor...";
        std::wstring dummy; std::getline(std::wcin, dummy);

        std::wstring uidHex, readerUsed;
        std::string err;
        if (!read_uid_pcsc_auto_wait(uidHex, readerUsed, err, 15000)) { std::wcerr << L"Erro UID: " << to_wstring_acp(err) << L"\n"; continue; }
        std::wcout << L"UID: " << uidHex << L" (leitor: " << readerUsed << L")\n";

        print_card_gdi(to_wstring_acp(mat), uidHex);
        std::string uid_nospace = remove_spaces_w_to_ascii(uidHex);
        out << mat << "," << uid_nospace << "\n"; out.flush();
    }
    out.close();
    std::wcout << L"Fim do processamento.\n";
    return 0;
}

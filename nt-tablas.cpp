#include <iostream>
#include <sstream>
#include <cstring>
#include <cmath>
#include <fstream>
#include <string>
#include <cctype>
#include <vector>
#include <regex>
#include <sstream>
#include <stack>
#include <chrono>
#include <thread>
#include <cstdlib>
#include <random>
#include <iomanip>
#include <map>
#include <filesystem>

// Windows API
#include <windows.h>
#include <tlhelp32.h>
#include <comdef.h>
#include <Wbemidl.h>

// Unix/Linux specific
#include <unistd.h>

// Other specific headers
#include <conio.h>  // Console I/O functions (Windows)

using namespace std;
typedef std::string String;
namespace fs = std::filesystem;

// Good for debugging
void stoplog(string message);

// Functions that return a string to change the background or foreground of the console text by using rgb
string rgbBackgrounds(int r, int g, int b);
string rgbs(int r, int g, int b);

// Function to move the cursor to a console position
void gotoxy(int x, int y);
string gotoxys(int x, int y);

// Some colors
const std::string ASCII_BOLD = "\x1b[1m"; // Bold text
const std::string ASCII_REVERSE = "\x1b[7m"; // Reverse (swap foreground and background colors)
const std::string ASCII_GREEN = "\x1b[32m";
const std::string ASCII_BLINK = "\x1b[5m"; // Blinking text
const std::string ASCII_RESET = "\x1b[0m\x1b[37m"; // Reset text to default color and style
const std::string ASCII_UNDERLINE = "\x1b[4m";  
const std::string ASCII_DIM = "\x1b[2m";  
const std::string ASCII_BG_GREEN = "\x1b[42m";
const std::string ASCII_BG_WHITE = "\x1b[47m";
const std::string ASCII_BLACK = "\x1b[30m";
const std::string ASCII_YELLOW = "\x1b[33m";

// Function to get console resolution
vector<int> getConsoleResolution() {
    CONSOLE_SCREEN_BUFFER_INFO csbi;
    vector<int> resolution(2, 0);
    GetConsoleScreenBufferInfo(GetStdHandle(STD_OUTPUT_HANDLE), &csbi);
    resolution[0] = csbi.srWindow.Right - csbi.srWindow.Left + 1;
    resolution[1] = csbi.srWindow.Bottom - csbi.srWindow.Top + 1;
    return resolution;
}

// Define grid size
int numRows = getConsoleResolution()[1] - 5;
int numCols = min(getConsoleResolution()[0] / 15 - 2, 26);

// Some variables for files
String filename = "";
String filecontent = "";

// Function to check if a file exists
inline bool exists_test0 (const std::string& name);

// Classes of elements in documents
// Cell class
class Cell {
public:
    String content;
    int x, y;

    Cell(String content = "", int x = 0, int y = 0) 
        : content(content), x(x), y(y) {}

    String getContent() { return content; }
    void setContent(String content) { this->content = content; }

    // Agregar métodos de acceso para x e y
    int getX() { return x; }
    int getY() { return y; }
};

// Table class
class Table {
public:
    std::vector<std::vector<Cell>> cells; // Matriz de celdas

    Table() = default;

    // Constructor para inicializar una tabla con un número específico de filas y columnas
    Table(int rows, int cols) {
        cells.resize(rows, std::vector<Cell>(cols));
    }

    // Obtener la matriz de celdas
    std::vector<std::vector<Cell>>& getCells() {
        return cells;
    }

    // Obtener una celda específica
    Cell& getCell(int row, int col) {
        return cells[row][col];
    }

    // Establecer el contenido de una celda específica
    void setCell(int row, int col, const String& content) {
        if (row >= 0 && row < cells.size() && col >= 0 && col < cells[row].size()) {
            cells[row][col].setContent(content);
        } else {
            std::cerr << "Error: Índices de celda fuera de límites." << std::endl;
        }
    }
   
    int getRowCount() const { return cells.size(); }
    int getColCount() const { return cells[0].size(); }

    vector<string> getPrinteableText() {
        vector<string> lines;
        const int COLUMN_WIDTH = 15;  // Reduced from 20 to fit better
        const int ROW_NUMBER_WIDTH = 4;
        
        // Add column letters header
        string header = string(ROW_NUMBER_WIDTH + 1, ' ');  // Space for row numbers
        for (int col = 0; col < getColCount(); ++col) {
            string colLetter = string(1, 'A' + col);
            string right(COLUMN_WIDTH - colLetter.length(), ' ');
            header += colLetter + right;
        }
        lines.push_back(header);

        // Add row numbers and cell content
        for (int row = 0; row < getRowCount(); ++row) {
            string line;
            // Add row number
            string rowNumber = to_string(row + 1);
            line += rowNumber + string(ROW_NUMBER_WIDTH - rowNumber.length() - 1, ' ') + "|";
            
            // Add cell content
            int lastNonEmptyCol = -1;
            bool isEmptyRow = true;
            for (int col = 0; col < getColCount(); ++col) {
            string content = getCell(row, col).getContent();
            // Truncate content if too long
            if (content.length() > COLUMN_WIDTH) {
                content = content.substr(0, COLUMN_WIDTH - 3) + "...";
            }
            line += content + string(COLUMN_WIDTH - content.length(), ' ');

            if (!content.empty()) {
                lastNonEmptyCol = col;
                isEmptyRow = false;
            }
            }

            // Check if the row is empty
            if (!isEmptyRow) {
            // Check if there are no more non-empty columns at right
            if (lastNonEmptyCol < getColCount() - 1) {
                for (int col = lastNonEmptyCol + 1; col < getColCount(); ++col) {
                line += string(COLUMN_WIDTH, ' ');
                }
            }
            lines.push_back(line);
            }
        }
        return lines;
    }
};

// Tab class
class Tab {
public:
    Table table;
    String name;

    Tab(Table table = Table(), String name = "") 
        : table(table), name(name) {}

    Table& getTable() { return table; }
    String getName() { return name; }
};

// Document class
class Document {
public:
    std::vector<Tab> tabs;
    String name;

    Document() = default;
    Document(std::vector<Tab> tabs, String name) 
        : tabs(tabs), name(name) {}

    std::vector<Tab>& getTabs() { return tabs; }
    void addTab(const Tab& tab) { tabs.push_back(tab); }

    String getName() { return name; }

    // Método para actualizar una celda
    void updateCell(int tabIndex, int row, int col, String content) {
        if (tabIndex >= 0 && tabIndex < tabs.size()) {
            Table& table = tabs[tabIndex].getTable();
            table.setCell(row, col, content);
        } else {
            std::cerr << "Error: Índice de pestaña no válido." << std::endl;
        }
    }
};

// A function to get the text that should be saved on the document
string getAllText(Document& doc);

// A function to create the initial tab
void createInitialTab(Document& doc) {
    // Crear una tabla con 10 filas y 10 columnas
    Table table(numRows, numCols);

    // Inicializar las celdas con coordenadas y contenido vacío
    for (int row = 0; row < table.getRowCount(); ++row) {
        for (int col = 0; col < table.getColCount(); ++col) {
            table.getCell(row, col) = Cell("", col, row);
        }
    }

    // Crear la pestaña
    Tab tab(table, "Sheet1");

    // Agregar la pestaña al documento
    doc.addTab(tab);
}

// Pre-definition of cool dialogs on the console
string inputDialog(int winx, int winy, int winlen, int winhei, string wintitle, vector<string> wincontent, string winbackcolor, string winforecolor);
void normaldialog(int winx, int winy, int winlen, int winhei, string wintitle, vector<string> wincontent, string winbackcolor, string winforecolor);
void errordialog(int winx, int winy, int winlen, int winhei, string wintitle, vector<string> wincontent, string winbackcolor, string winforecolor);
int selectiondialog(int winx, int winy, int winlen, int winhei, string wintitle, vector<string> wincontent, vector<string> options, string winbackcolor, string winforecolor);
int bigselectiondialog(int winx, int winy, int winlen, int winhei, string wintitle, vector<string> wincontent, vector<string> options, string winbackcolor, string winforecolor);

// A function to print documents in a printer
void print(const vector<string>& lines) {
    // No way this works

    PRINTER_INFO_2* pPrinterInfo = nullptr;
    DWORD dwNeeded, dwReturned;

    if (!EnumPrinters(PRINTER_ENUM_LOCAL, nullptr, 2, nullptr, 0, &dwNeeded, &dwReturned)) {
        if (GetLastError() == ERROR_INSUFFICIENT_BUFFER) {
            pPrinterInfo = (PRINTER_INFO_2*)malloc(dwNeeded);
            if (pPrinterInfo != nullptr) {
                if (EnumPrinters(PRINTER_ENUM_LOCAL, nullptr, 2, (LPBYTE)pPrinterInfo, dwNeeded, &dwNeeded, &dwReturned)) {
                    vector<string> printers;
                    for (DWORD i = 0; i < dwReturned; ++i) {
                        printers.push_back(string(pPrinterInfo[i].pPrinterName));
                    }
                    
                    string printer = printers[bigselectiondialog(getConsoleResolution()[0] / 2 - 25, getConsoleResolution()[1] / 2 - 25 / 2, 50, 25, "eLite NT Office 2025", {"Select a printer: "}, printers, rgbBackgrounds(240,240,240), ASCII_BLACK)];


                    // Printer setup with landscape mode
                    DEVMODE* pDevMode = nullptr;
                    HANDLE hPrinter;
                    if (OpenPrinter(const_cast<char*>(printer.c_str()), &hPrinter, NULL)) {
                        DWORD dwNeededDevMode = DocumentProperties(NULL, hPrinter, 
                            const_cast<char*>(printer.c_str()), NULL, NULL, 0);
                        pDevMode = (DEVMODE*)malloc(dwNeededDevMode);
                        if (pDevMode && DocumentProperties(NULL, hPrinter, 
                            const_cast<char*>(printer.c_str()), pDevMode, NULL, DM_OUT_BUFFER) == IDOK) {
                            pDevMode->dmOrientation = DMORIENT_LANDSCAPE;
                            pDevMode->dmFields |= DM_ORIENTATION;
                        }
                        ClosePrinter(hPrinter);
                    }

                    HDC printerHandle = CreateDC(NULL, printer.c_str(), NULL, pDevMode);
                    if (!printerHandle) {
                        // Error handling
                        if (pDevMode) free(pDevMode);
                        return;
                    }

                    // Get printer metrics
                    int dpiX = GetDeviceCaps(printerHandle, LOGPIXELSX);
                    int dpiY = GetDeviceCaps(printerHandle, LOGPIXELSY);
                    int margin = static_cast<int>(0.05 * dpiX); // 0.5cm margin
                    int fontSize = static_cast<int>(12 * dpiY / 72.0); // 12pt font

                    DOCINFO docInfo = { sizeof(DOCINFO), filename.c_str(), NULL, NULL, 0 };
                    if (StartDoc(printerHandle, &docInfo) > 0 && StartPage(printerHandle) > 0) {
                        // Create font with underline
                        HFONT hFont = CreateFont(
                            -fontSize, 0, 0, 0, FW_NORMAL, FALSE, TRUE, FALSE,
                            DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                            DEFAULT_QUALITY, FIXED_PITCH | FF_MODERN, "Consolas");
                        if (hFont) {
                            SelectObject(printerHandle, hFont);
                            
                            // Get text metrics for proper spacing
                            TEXTMETRIC tm;
                            GetTextMetrics(printerHandle, &tm);
                            int lineHeight = tm.tmHeight + tm.tmExternalLeading;
                            int yPos = margin;

                            // Print each line
                            for (const auto& line : lines) {
                                if (yPos + lineHeight > GetDeviceCaps(printerHandle, VERTRES) - margin) {
                                    // Start new page if we're at the bottom margin
                                    EndPage(printerHandle);
                                    StartPage(printerHandle);
                                    yPos = margin;
                                }
                                TextOutA(printerHandle, margin, yPos, line.c_str(), line.length());
                                yPos += lineHeight;
                            }
                            DeleteObject(hFont);
                        }
                        EndPage(printerHandle);
                        EndDoc(printerHandle);
                    }
                    DeleteDC(printerHandle);
                    if (pDevMode) free(pDevMode);
                }
                free(pPrinterInfo);
            }
        }
    }
}

// Definition of cool in-console dialogs
string inputDialog(int winx, int winy, int winlen, int winhei, string wintitle, vector<string> wincontent, string winbackcolor, string winforecolor) {
    string winbackground(winlen, ' ');
    
    // Draw the window background
    stringstream ss;
    for (int i = 0; i < winhei; i++) {
        gotoxy(winx, winy + i);
        ss << gotoxys(winx, winy + i) << winbackcolor << winbackground << ASCII_RESET;
    }
    cout << ss.str();

    string bottom(winlen, 223);
    gotoxy(winx + 1, winy + winhei);
    cout << ASCII_DIM << rgbs(255,255,255) << bottom << ASCII_RESET;

    for(int i = 0; i < winhei - 1; i++) {
        gotoxy(winx + winlen, winy + i + 1);
        cout << ASCII_DIM << rgbs(255,255,255) << (char) 219 << ASCII_RESET;
    }

    // Draw the title bar
    gotoxy(winx, winy);
    string titlebar(winlen - 1 - wintitle.length(), ' ');
    cout << rgbBackgrounds(21, 82, 26) << " " << ASCII_BOLD << wintitle << titlebar << ASCII_RESET;

    // Draw the window content
    for(int i = 0; i < wincontent.size(); i++) {
        gotoxy(winx + 1, winy + 2 + i);
        cout << winbackcolor << winforecolor << wincontent[i] << ASCII_RESET;
    }


    // Position cursor for user input
    gotoxy(winx + 1, winy + winhei - 2);
    string inputbar(winlen - 2, ' ');
    cout << ASCII_BG_WHITE << ASCII_BLACK << inputbar << ASCII_BOLD;
    gotoxy(winx + 1, winy + winhei - 2);
    string result;
    getline(cin, result);

    cout << ASCII_RESET;
    return result;
}

void normaldialog(int winx, int winy, int winlen, int winhei, string wintitle, vector<string> wincontent, string winbackcolor, string winforecolor) {
    string winbackground(winlen, ' ');

    // Draw the window background
    stringstream ss;
    for (int i = 0; i < winhei; i++) {
        gotoxy(winx, winy + i);
        ss << gotoxys(winx, winy + i) << winbackcolor << winbackground << ASCII_RESET;
    }
    cout << ss.str();

    string bottom(winlen, 223);
    gotoxy(winx + 1, winy + winhei);
    cout << ASCII_DIM << rgbs(255,255,255) << bottom << ASCII_RESET;

    for(int i = 0; i < winhei - 1; i++) {
        gotoxy(winx + winlen, winy + i + 1);
        cout << ASCII_DIM << rgbs(255,255,255) << (char) 219 << ASCII_RESET;
    }

    gotoxy(winx+winlen, winy);
    cout << ASCII_DIM << rgbs(255,255,255) << (char) 220 << ASCII_RESET;

    // Draw the title bar
    gotoxy(winx, winy);
    string titlebar(winlen - 1 - wintitle.length(), ' ');
    cout << rgbBackgrounds(21, 82, 26) << " " << ASCII_BOLD << wintitle << titlebar << ASCII_RESET;

    // Draw the window content
    for(int i = 0; i < wincontent.size(); i++) {
        gotoxy(winx + 1, winy + 2 + i);
        cout << winbackcolor << winforecolor << wincontent[i] << ASCII_RESET;
    }
    
    // Draw ok button
    string okbutton = " Ok ";
    string okbuttonbottom(okbutton.length(), 223);
    gotoxy(winx + winlen / 2 - okbutton.length() / 2, winy + winhei - 3);
    cout << rgbBackgrounds(255,255,255) << ASCII_BLACK << ASCII_BOLD << okbutton << ASCII_RESET << ASCII_DIM << rgbs(100,100,100) << winbackcolor << (char) 220 << ASCII_RESET;
    gotoxy(winx + winlen / 2 - okbutton.length() / 2 + 1, winy + winhei - 2);
    cout << winbackcolor << ASCII_DIM << rgbs(100,100,100) << okbuttonbottom << ASCII_RESET;
    _getch();

    cout << ASCII_RESET;
}

void errordialog(int winx, int winy, int winlen, int winhei, string wintitle, vector<string> wincontent, string winbackcolor, string winforecolor) {
    string winbackground(winlen, ' ');

    // Draw the window background
    stringstream ss;
    for (int i = 0; i < winhei; i++) {
        gotoxy(winx, winy + i);
        ss << gotoxys(winx, winy + i) << winbackcolor << winbackground << ASCII_RESET;
    }
    cout << ss.str();

    string bottom(winlen, 223);
    gotoxy(winx + 1, winy + winhei);
    cout << ASCII_DIM << rgbs(255,255,255) << bottom << ASCII_RESET;

    for(int i = 0; i < winhei - 1; i++) {
        gotoxy(winx + winlen, winy + i + 1);
        cout << ASCII_DIM << rgbs(255,255,255) << (char) 219 << ASCII_RESET;
    }

    gotoxy(winx+winlen, winy);
    cout << ASCII_DIM << rgbs(255,255,255) << (char) 220 << ASCII_RESET;

    // Draw the title bar
    gotoxy(winx, winy);
    string titlebar(winlen - 1 - wintitle.length(), ' ');
    cout << rgbBackgrounds(235, 64, 52) << " " << ASCII_BOLD << wintitle << titlebar << ASCII_RESET;

    // Draw the window content
    for(int i = 0; i < wincontent.size(); i++) {
        gotoxy(winx + 1, winy + 2 + i);
        cout << winbackcolor << winforecolor << wincontent[i] << ASCII_RESET;
    }
    
    // Draw ok button
    string okbutton = " Ok ";
    string okbuttonbottom(okbutton.length(), 223);
    gotoxy(winx + winlen / 2 - okbutton.length() / 2, winy + winhei - 2);
    cout << rgbBackgrounds(255,255,255) << ASCII_BLACK << ASCII_BOLD << okbutton << ASCII_RESET << ASCII_DIM << rgbs(100,100,100) << winbackcolor << (char) 220 << ASCII_RESET;
    gotoxy(winx + winlen / 2 - okbutton.length() / 2 + 1, winy + winhei - 1);
    cout << winbackcolor << ASCII_DIM << rgbs(100,100,100) << okbuttonbottom << ASCII_RESET;
    
    cout << "\a";

    _getch();

    cout << ASCII_RESET;
}

int selectiondialog(int winx, int winy, int winlen, int winhei, string wintitle, vector<string> wincontent, vector<string> options, string winbackcolor, string winforecolor)  {
    string winbackground(winlen, ' ');

    // Draw the window background
    stringstream ss;
    for (int i = 0; i < winhei; i++) {
        gotoxy(winx, winy + i);
        ss << gotoxys(winx, winy + i) << winbackcolor << winbackground << ASCII_RESET;
    }
    cout << ss.str();

    string bottom(winlen, 223);
    gotoxy(winx + 1, winy + winhei);
    cout << ASCII_DIM << rgbs(255,255,255) << bottom << ASCII_RESET;

    for(int i = 0; i < winhei - 1; i++) {
        gotoxy(winx + winlen, winy + i + 1);
        cout << ASCII_DIM << rgbs(255,255,255) << (char) 219 << ASCII_RESET;
    }

    gotoxy(winx+winlen, winy);
    cout << ASCII_DIM << rgbs(255,255,255) << (char) 220 << ASCII_RESET;

    // Draw the title bar
    gotoxy(winx, winy);
    string titlebar(winlen - 1 - wintitle.length(), ' ');
    cout << rgbBackgrounds(21, 82, 26) << " " << ASCII_BOLD << wintitle << titlebar << ASCII_RESET;

    // Draw the window content
    for(int i = 0; i < wincontent.size(); i++) {
        gotoxy(winx + 1, winy + 2 + i);
        cout << winbackcolor << winforecolor << wincontent[i] << ASCII_RESET;
    }

    int selection = 0;
    bool selected = false;

    // Draw the options
    while(!selected) {
        stringstream ss2;
        int optionY = winy + winhei - 2;
        int optionslen = 0;
        for(int i = 0; i < options.size(); i++) {
            string option = options[i];
            string optionbottom(option.length() + 2, 223);
            ss2 << gotoxys(winx + 1 + optionslen, optionY) << rgbBackgrounds(255,255,255) << ASCII_BLACK;
            if(selection == i) { ss2 << ASCII_BOLD; }
            else { ss2 << ASCII_DIM; }
            ss2 << " " + option + " " << ASCII_RESET << ASCII_DIM << rgbs(100,100,100) << winbackcolor << (char) 220 << ASCII_RESET;
            ss2 << gotoxys(winx + 2 + optionslen, optionY + 1) << winbackcolor << ASCII_DIM << rgbs(100,100,100) << optionbottom << ASCII_RESET;
            optionslen += option.length() + 4;
        }
        cout << ss2.str();
        char key = _getch();
        if(key == 75) {
            selection = max(0, selection - 1);
        } else if(key == 77) {
            selection = min((int) options.size() - 1, selection + 1);
        } else if(key == 13) {
            selected = true;
        }
    }

    return selection;
}

int bigselectiondialog(int winx, int winy, int winlen, int winhei, string wintitle, vector<string> wincontent, vector<string> options, string winbackcolor, string winforecolor) {
    string winbackground(winlen, ' ');

    // Draw the window background
    stringstream ss;
    for (int i = 0; i < winhei; i++) {
        gotoxy(winx, winy + i);
        ss << gotoxys(winx, winy + i) << winbackcolor << winbackground << ASCII_RESET;
    }
    cout << ss.str();

    string bottom(winlen, 223);
    gotoxy(winx + 1, winy + winhei);
    cout << ASCII_DIM << rgbs(255,255,255) << bottom << ASCII_RESET;

    for(int i = 0; i < winhei - 1; i++) {
        gotoxy(winx + winlen, winy + i + 1);
        cout << ASCII_DIM << rgbs(255,255,255) << (char) 219 << ASCII_RESET;
    }

    gotoxy(winx+winlen, winy);
    cout << ASCII_DIM << rgbs(255,255,255) << (char) 220 << ASCII_RESET;

    // Draw the title bar
    gotoxy(winx, winy);
    string titlebar(winlen - 1 - wintitle.length(), ' ');
    cout << rgbBackgrounds(21, 82, 26) << " " << ASCII_BOLD << wintitle << titlebar << ASCII_RESET;

    // Draw the window content
    for(int i = 0; i < wincontent.size(); i++) {
        gotoxy(winx + 1, winy + 2 + i);
        cout << winbackcolor << winforecolor << wincontent[i] << ASCII_RESET;
    }

    int selection = 0;
    bool selected = false;

    // Draw the options
    while(!selected) {
        stringstream ss2;
        int sopy = winy + 2 + wincontent.size();
        int listwisth = winlen - 2;
        for(int i = 0; i < options.size(); i++) {
            string option = options[i];

            if(i == selection) {
                ss2 << ASCII_BOLD;
            } else {
                ss2 << ASCII_DIM;
            }

            string leftbar(listwisth - option.length(), ' ');
            ss2 << gotoxys(winx + 1, sopy + i) << rgbBackgrounds(200,200,200) << ASCII_BLACK << option << leftbar << ASCII_RESET;
        }
        cout << ss2.str();
        char key = _getch();
        if(key == 72) { // Up arrow key
            selection = max(0, selection - 1);
        } else if(key == 80) { // Down arrow key
            selection = min((int) options.size() - 1, selection + 1);
        } else if(key == 13) { // Enter key
            selected = true;
        }
    }

    return selection;
}

// Variable which stores the opened tab
int openedTab = 0;

// A function to save the document
void saveDocument(Document& doc) {
    if(doc.getName().empty() || doc.getName() == "Untitled") {
        // Show the dialog to select a file for saving it
        int winx = getConsoleResolution()[0] / 2 - 70 / 2;
        int winy = getConsoleResolution()[1] / 2 - 6;

        int winlen = 70;
        int winhei = 6;

        string wintitle = "Save document";
        string wincontent = "No document name specified. Please enter a name for the document: ";

        string winbackcolor = rgbBackgrounds(240, 240, 240);
        string winforecolor = rgbs(0,0,0);
        
        string winbackground(winlen, ' ');

        string filename = inputDialog(winx, winy, winlen, winhei, wintitle, {wincontent}, winbackcolor, winforecolor);
        doc.name = filename;
        saveDocument(doc);
    } else {
        // The document has a name, so we can save it
        ofstream outfile(doc.getName());
        outfile << getAllText(doc);
        outfile.close();
    }

    // Save the document
    ofstream outfile(doc.getName());
    if (outfile.good()) {
        outfile << getAllText(doc);
        outfile.close();
    } else {
        std::cerr << "Error: No se pudo abrir el archivo para guardar." << std::endl;
    }
}

// A function to save the document as a new one
void saveasDocument(Document& doc) {
    // Just a dialog asking for the name of the document
    int winx = getConsoleResolution()[0] / 2 - 70 / 2;
    int winy = getConsoleResolution()[1] / 2 - 6;

    int winlen = 70;
    int winhei = 6;

    string wintitle = "Save as document";
    string wincontent = "Enter a name for the document: ";

    string winbackcolor = rgbBackgrounds(240, 240, 240);
    string winforecolor = rgbs(0,0,0);
    
    string winbackground(winlen, ' ');

    string filename = inputDialog(winx, winy, winlen, winhei, wintitle, {wincontent}, winbackcolor, winforecolor);

    Document newdoc = Document(doc.getTabs(), filename);

    ofstream outfile(newdoc.getName());
    if (outfile.good()) {
        outfile << getAllText(newdoc);
        outfile.close();
    } else {
        std::cerr << "Error: No se pudo abrir el archivo para guardar." << std::endl;
    }
}

// A function that returns a document from a file
Document openDocument(string filename2, Document openedDocument) {  
    if(exists_test0(filename2)) {
        ifstream infile(filename2);
        vector<Tab> tabs;
        string line;

        // Leer el archivo y cargar las celdas
        while (getline(infile, line)) {
            if (line.at(0) == 's') {
                // Starting writing tab
                string tabname = line.substr(1);
                Table table(numRows, numCols);
                tabs.push_back(Tab(table, tabname));
            } else {
                istringstream iss(line);
                string token;
                vector<string> cellInfo;
                while (getline(iss, token, ';')) {
                    cellInfo.push_back(token);
                }
                if (cellInfo.size() == 3) {
                    int x = stoi(cellInfo[0]);
                    int y = stoi(cellInfo[1]);
                    string content = cellInfo[2];

                    // Ensure the matrix has enough space
                    if (y >= tabs.back().getTable().getRowCount()) {
                        tabs.back().getTable().cells.resize(y + 1);
                    }
                    if (x >= tabs.back().getTable().cells[y].size()) {
                        tabs.back().getTable().cells[y].resize(x + 1);
                    }

                    // Add the cell to the matrix
                    tabs.back().getTable().setCell(y, x, content);
                }
            }
        }

        // Devolver el documento cargado
        return Document(tabs, filename2);
    } else {
        errordialog(getConsoleResolution()[0] / 2 - 50 / 2, getConsoleResolution()[1] / 2 - 10 / 2, 50, 10, "Error", {"Specified file doesn't exist,", "Check for spelling errors."}, rgbBackgrounds(240,240,240), ASCII_BLACK);
    }
    return openedDocument;
}

// Variables storing the podition of the mouse in the table
int cursorX = 0, cursorY = 0;

// Define grid data structure
vector<vector<string>> grid(numRows, vector<string>(numCols, ""));

// Ojito q va dpm
int calculate(const std::string& expression) {
    std::stack<int> numbers;
    std::stack<char> operators;

    std::istringstream iss(expression);

    char currentChar;
    while (iss >> currentChar) {
        if (isdigit(currentChar)) {
            iss.putback(currentChar);
            int number;
            iss >> number;
            numbers.push(number);
        } else if (currentChar == '(') {
            operators.push(currentChar);
        } else if (currentChar == ')') {
            while (!operators.empty() && operators.top() != '(') {
                int operand2 = numbers.top();
                numbers.pop();
                int operand1 = numbers.top();
                numbers.pop();
                char op = operators.top();
                operators.pop();

                if (op == '+') {
                    numbers.push(operand1 + operand2);
                } else if (op == '-') {
                    numbers.push(operand1 - operand2);
                } else if (op == '*') {
                    numbers.push(operand1 * operand2);
                } else if (op == '/') {
                    numbers.push(operand1 / operand2);
                }
            }
            operators.pop();  // Pop '('
        } else if (currentChar == '+' || currentChar == '-' || currentChar == '*' || currentChar == '/') {
            while (!operators.empty() && (operators.top() == '*' || operators.top() == '/') && currentChar != '(') {
                int operand2 = numbers.top();
                numbers.pop();
                int operand1 = numbers.top();
                numbers.pop();
                char op = operators.top();
                operators.pop();

                if (op == '+') {
                    numbers.push(operand1 + operand2);
                } else if (op == '-') {
                    numbers.push(operand1 - operand2);
                } else if (op == '*') {
                    numbers.push(operand1 * operand2);
                } else if (op == '/') {
                    numbers.push(operand1 / operand2);
                }
            }
            operators.push(currentChar);
        }
    }

    while (!operators.empty()) {
        int operand2 = numbers.top();
        numbers.pop();
        int operand1 = numbers.top();
        numbers.pop();
        char op = operators.top();
        operators.pop();

        if (op == '+') {
            numbers.push(operand1 + operand2);
        } else if (op == '-') {
            numbers.push(operand1 - operand2);
        } else if (op == '*') {
            numbers.push(operand1 * operand2);
        } else if (op == '/') {
            numbers.push(operand1 / operand2);
        }
    }

    return numbers.top();
}

string cellToString(int row, int col, const string& content) {
    return to_string(col) + ";" + to_string(row) + ";" + content;
}

// Function to generate a multiline string containing cell information
string getAllText(Document& doc) {
    string result;

    for (int i = 0; i < doc.getTabs().size(); i++) {
        result += "s" + doc.getTabs()[i].getName() + "\n";
        Table& table = doc.getTabs()[i].getTable();
        for (int row = 0; row < table.getRowCount(); ++row) {
            for (int col = 0; col < table.getColCount(); ++col) {
                Cell& cell = table.getCell(row, col);
                if (!cell.getContent().empty()) {
                    result += cellToString(row, col, cell.getContent()) + "\n";
                }
            }
        }
    }
    
    return result;
}
// Function to split an array of strings into smaller arrays of a determined length
vector<vector<string>> splitArray(const string comms[], int arraySize, int subArrayLength) {
    vector<vector<string>> result;

    for (int i = 0; i < arraySize; i += subArrayLength) {
        vector<string> subArray;

        // Copy elements from comms[] to subArray
        for (int j = i; j < min(i + subArrayLength, arraySize); ++j) {
            subArray.push_back(comms[j]);
        }

        // Add subArray to result
        result.push_back(subArray);
    }

    return result;
}


// Function to set cursor position
void gotoxy(int x, int y) {
    COORD coord;
    coord.X = x;
    coord.Y = y;
    SetConsoleCursorPosition(GetStdHandle(STD_OUTPUT_HANDLE), coord);
}
// Function that resturns in a string how to move the cursor
string gotoxys(int x, int y) {
    return "\033[" + to_string(y + 1) + ";" + to_string(x + 1) + "H";
}

// Function to extract row and column indices from cell reference
pair<int, int> parseCellReference(const string& cellRef) {
    int row, col;
    try {
        if (isdigit(cellRef[0])) {
            row = stoi(cellRef) - 1;
            col = 0;
        } else {
            col = cellRef[0] - 'A';
            row = stoi(cellRef.substr(1)) - 1;
        }
    } catch (const std::invalid_argument& e) {
        // Handle invalid input gracefully
        row = -1; // Invalid row
        col = -1; // Invalid column
    }
    return {row, col};
}
std::string getTextInLastBraquets(const std::string& str) {
    int openBracketPos = -1;
    int closeBracketPos = -1;
    int bracketCount = 0;
    bool isInsideQuotes = false;

    for (int i = 0; i < str.length(); ++i) {
        if (str[i] == '"') {
            isInsideQuotes = !isInsideQuotes;
        }

        if (!isInsideQuotes) {
            if (str[i] == '(') {
                if (bracketCount == 0) {
                    openBracketPos = i;
                }
                bracketCount++;
            } else if (str[i] == ')') {
                bracketCount--;
                if (bracketCount == 0) {
                    closeBracketPos = i;
                }
            }
        }
    }

    if (openBracketPos != -1 && closeBracketPos != -1) {
        return str.substr(openBracketPos + 1, closeBracketPos - openBracketPos - 1);
    } else {
        return ""; // No complete bracket pair found
    }
}

// Function to replace expressions inside CALC formulas with their results
void replaceCExpressions(string& expression) {
    size_t pos = expression.find("c(");
    while (pos != string::npos) {
        size_t endPos = expression.find(")", pos);
        if (endPos != string::npos) {
            string subExpression = expression.substr(pos + 2, endPos - pos - 2);
            int result = calculate(subExpression);
            expression.replace(pos, endPos - pos + 1, to_string(result));
        }
        pos = expression.find("c(", pos + 1); // Start searching from the next character
    }
}

bool isNum(string str) {
    for(int i = 0; i < str.length(); i++) {
        if(str.c_str()[i] != '0' && str.c_str()[i] != '1' && str.c_str()[i] != '2' && str.c_str()[i] != '3' && str.c_str()[i] != '4' && str.c_str()[i] != '5' && str.c_str()[i] != '6' && str.c_str()[i] != '7' && str.c_str()[i] != '8' && str.c_str()[i] != '9' && str.c_str()[i] != '+' && str.c_str()[i] != '-' && str.c_str()[i] != '/' && str.c_str()[i] != '*') return false;
    }
    return true;
}
string evaluateFormula(const string& formula, int posx, int posy, Document openedDocument);

string getCellValue(const string& cellRef, int posx, int posy, Document& document, int numCols) {
    // Convertir la referencia de celda (por ejemplo, "A1") a índices de fila y columna
    int col = cellRef[0] - 'A';
    int row = stoi(cellRef.substr(1)) - 1;

    // Verificar que los índices estén dentro de los límites
    if (row >= 0 && row < document.getTabs()[0].getTable().getRowCount() &&
        col >= 0 && col < document.getTabs()[0].getTable().getColCount()) {
        // Obtener el contenido de la celda
        return document.getTabs()[0].getTable().getCell(row, col).getContent();
    } else {
        return "#REF"; // Referencia inválida
    }
}

// Function to evaluate a CONCAT formula
string evaluateCONCATFormula(const string& formula, int posx, int posy, Document openedDocument) {
    // Extract the text inside the CONCAT formula
    string text = getTextInLastBraquets(formula);

    // Check if the text is empty
    if (text.empty()) {
        return rgbs(235, 64, 52) + ASCII_BOLD + "#FORMU" + ASCII_RESET; // Return #FORMU for unknown formula
    }

    // Split the text into arguments using ';' as delimiter
    vector<string> arguments;
    size_t startPos = 0, endPos;
    while ((endPos = text.find(';', startPos)) != string::npos) {
        arguments.push_back(text.substr(startPos, endPos - startPos));
        startPos = endPos + 1;
    }
    arguments.push_back(text.substr(startPos)); // Add the last argument

    // Evaluate each argument and concatenate them together
    string concatenatedText;
    for (const string& arg : arguments) {
        string evaluatedArg = evaluateFormula(arg, posx, posy, openedDocument); // Evaluate the argument
        concatenatedText += evaluatedArg;
    }

    return concatenatedText;
}

void replaceAllOutsideQuotes(std::string& str, const std::string& from, const std::string& to) {
    size_t start_pos = 0;
    bool inside_quotes = false;
    
    while ((start_pos = str.find(from, start_pos)) != std::string::npos) {
        // Check if the occurrence of 'from' is inside quotes
        size_t quote_start = str.rfind('"', start_pos);
        size_t quote_end = str.find('"', start_pos);
        
        // If both quotes are found and the occurrence of 'from' is within them, skip replacement
        if (quote_start != std::string::npos && quote_end != std::string::npos && quote_start < start_pos && quote_end > start_pos) {
            start_pos = quote_end + 1; // Move to the character after the end quote
            continue;
        }
        
        // Replace occurrence of 'from' outside quotes
        str.replace(start_pos, from.length(), to);
        start_pos += to.length(); // Handles case where 'to' is a substring of 'from'
    }
}
// Función para evaluar una fórmula
string evaluateFormula(const string& formula, int posx, int posy, Document openedDocument) {
    if (formula.empty() || formula[0] != '=') {
        return formula; // No es una fórmula
    }

    string text = getTextInLastBraquets(formula);

    if (formula.substr(0, 5) == "=CALC") {
        string calcFormula = text;

        // Attempt to calculate the result
        int result;
        if (isNum(calcFormula)) {
            // If the formula contains only numeric characters, directly calculate the result
            result = calculate(calcFormula);
        } else {
            // Otherwise, replace cell references with their values and calculate the result
            size_t pos = 0;
            while ((pos = calcFormula.find_first_of("ABCDEFGHIJKLMNOPQRSTUVWXYZ", pos)) != string::npos) {
                // Extract the cell reference
                string cellRef;
                size_t endPos = calcFormula.find_first_not_of("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789", pos);
                if (endPos != string::npos) {
                    cellRef = calcFormula.substr(pos, endPos - pos);
                } else {
                    cellRef = calcFormula.substr(pos);
                }

                // Get the value of the cell reference and replace it in the formula
                string cellValue = getCellValue(cellRef, posx, posy, openedDocument, openedTab);
                if (cellValue != "#REF") {
                    calcFormula.replace(pos, cellRef.length(), cellValue);
                    // Update pos to search for the next occurrence
                    pos += cellValue.length();
                } else {
                    // Handle reference error
                    return "#REF";
                }
            }

            // Calculate the result
            result = calculate(calcFormula);
        }

        // Return the calculated result
        return to_string(result);
    } else if (formula.substr(0, 4) == "=GET") {
        string cellRef = formula.substr(5); // Extraer la referencia (por ejemplo, "A1")
        pair<int, int> cellPos = parseCellReference(cellRef);
        if (cellPos.first >= 0 && cellPos.first < numRows && cellPos.second >= 0 && cellPos.second < numCols) {
            return getCellValue(cellRef, posx, posy, openedDocument, openedDocument.getTabs()[0].getTable().getColCount());
        } else {
            return "#REF"; // Referencia inválida
        }
    }
    return formula;
}

void replaceAll(std::string& str, const std::string& from, const std::string& to) {
    size_t start_pos = 0;
    while ((start_pos = str.find(from, start_pos)) != std::string::npos) {
        str.replace(start_pos, from.length(), to);
        start_pos += to.length(); // Handles case where 'to' is a substring of 'from'
    }
}
void rgb(int r, int g, int b) {
    cout <<  "\033[38;2;" + to_string(r) + ";" + to_string(g) + ";" + to_string(b) + "m";
}
void rgbBackground(int r, int g, int b) {
    // ANSI escape code for setting background color using RGB
    cout << "\033[48;2;" + to_string(r) + ";" + to_string(g) + ";" + to_string(b) + "m";
}
string rgbs(int r, int g, int b) {
    HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE);
    CONSOLE_SCREEN_BUFFER_INFO consoleInfo;
    GetConsoleScreenBufferInfo(hConsole, &consoleInfo);
    WORD originalAttributes = consoleInfo.wAttributes;
    return "\033[38;2;" + to_string(r) + ";" + to_string(g) + ";" + to_string(b) + "m";
}
string rgbBackgrounds(int r, int g, int b) {
    // ANSI escape code for setting background color using RGB
    return "\033[48;2;" + to_string(r) + ";" + to_string(g) + ";" + to_string(b) + "m";
}

String getCellContent(Document* openedDocument, int openedTab, int cursorY, int cursorX, int numCols) {
    if (!openedDocument || openedTab < 0 || openedTab >= openedDocument->getTabs().size()) {
        return ""; // Documento o pestaña inválida
    }

    Table& table = openedDocument->getTabs()[openedTab].getTable();
    if (cursorY >= 0 && cursorY < table.getRowCount() && cursorX >= 0 && cursorX < table.getColCount()) {
        return table.getCell(cursorY, cursorX).getContent();
    } else {
        return ""; // Índices fuera de límites
    }
}

void drawGrid(int cursorX, int cursorY, Document& openedDocument) {
    // Limpiar la consola
    gotoxy(0,0);

    // Verificar límites del cursor
    cursorX = max(0, min(cursorX, numCols - 1));
    cursorY = max(0, min(cursorY, numRows - 1));

    // Obtener el nombre del documento
    String name = openedDocument.getName();
    string title = " eLite NT Office 2025 Tablas - " + name + " | Cursor: " + char('A' + cursorX) + to_string(cursorY + 1) + " ";

    // Dibujar el título
    gotoxy(0, 0);
    string tbar = "";
    for (int i = 0; i < getConsoleResolution()[0] / 2 - title.length() / 2; i++) {
        tbar += " ";
    }
    cout << ASCII_DIM << ASCII_GREEN << ASCII_REVERSE << tbar << ASCII_RESET;
    cout << ASCII_GREEN << ASCII_REVERSE << title << ASCII_RESET;
    tbar = "";
    for (int i = 0; i < getConsoleResolution()[0] / 2 - title.length() / 2 - 1; i++) {
        tbar += " ";
    }
    cout << ASCII_DIM << ASCII_GREEN << ASCII_REVERSE << tbar << ASCII_RESET;
    cout << endl;

    // Mostrar el contenido de la celda seleccionada
    String txt = " Text: " + ASCII_BOLD + openedDocument.getTabs()[openedTab].getTable().getCell(cursorY, cursorX).getContent();
    cout << rgbBackgrounds(240,240,240) << ASCII_BLACK << txt;
    string bar(" ", getConsoleResolution()[0] - txt.length());
    cout << setw(getConsoleResolution()[0] - txt.length()) << left << " " << ASCII_RESET;
    string tablist;
    for(int i = 0; i < openedDocument.getTabs().size(); i++) {
        if(i == openedTab) {
            tablist += rgbBackgrounds(240,240,240) + ASCII_BLACK + " "  + openedDocument.getTabs()[i].getName() + " " + ASCII_RESET;
        } else {
            tablist += rgbBackgrounds(240,240,240) + rgbs(33, 33, 33) + " " + openedDocument.getTabs()[i].getName() + " ";
        }
    }

    string space(getConsoleResolution()[0], ' ');
    stringstream menubar;

    menubar << "\n" << rgbBackgrounds(210,210,210) << ASCII_BLACK << " Menu " << ASCII_RESET << rgbBackgrounds(240,240,240) << rgbs(33, 33, 33) << " Tabs: " << tablist << rgbBackgrounds(240,240,240) + space;

    cout << menubar.str() <<  "\n\n" << ASCII_RESET;

    // Dibujar las letras de las columnas
    gotoxy(0, 3);
    cout << "    |";
    cout << ASCII_UNDERLINE;
    stringstream ss;
    for (int i = 0; i < numCols; ++i) {
        if (i < 26) { // Only show letters A to Z
            char letter = 'A' + i;
            ss << ASCII_UNDERLINE << " ";

            if(i == cursorX) {
                ss << ASCII_BOLD << letter; // Adjust spacing as needed
            } else {
                ss << letter; // Adjust spacing as needed
            }
            ss << ASCII_RESET << ASCII_UNDERLINE << "             |";
        } else {
            break; // Do not show more columns if they exceed Z
        }
    }
    ss << endl << ASCII_RESET;
    cout << ss.str();

    // Draw the grid
    for (int i = 0; i < numRows; ++i) {
        stringstream ss;

        // Draw the row number
        ss << ASCII_UNDERLINE;

        if(i == cursorY) {
            ss << ASCII_BOLD << setw(3) << left << i + 1 << " |" << ASCII_RESET;
        } else {
            ss << setw(3) << left << i + 1 << " |";
        }

        // Draw the cells
        for (int j = 0; j < numCols; ++j) {
            // Highlight the selected cell
            ss << ASCII_UNDERLINE;
            if (i == cursorY && j == cursorX) {
                ss << "\033[48;2;16;93;201m"; // rgbBackground(16, 93, 201);
            }

            Table table = openedDocument.getTabs()[openedTab].getTable();
            Cell cell = table.getCell(i, j);
            string cellContent = cell.getContent();

            // Evaluate formulas if necessary
            string evaluatedContent = evaluateFormula(cellContent, j, i, openedDocument);

            // Display the cell content (only the first 10 characters)
            if (evaluatedContent.empty()) {
                ss << setw(15) << left << ""; // Leave empty cells as they are
            } else if (evaluatedContent.length() >= 15) {
                ss << evaluatedContent.substr(0, 7) << "...";
            } else {
                ss << setw(15) << left << evaluatedContent;
            }

            // Reset the format
            ss << ASCII_RESET << "|";
        }
        ss << endl;
        cout << ss.str();
    }
}

String openedFolder;

inline bool exists_test0 (const std::string& name) {
    ifstream f(name.c_str());
    return f.good();
}


String getFileState(Document& doc) {
    if(exists_test0(doc.getName())) {
        fstream f(doc.getName());
        String l, savedFileTxt;
        while(getline(f, l)) {
            savedFileTxt += l + "\n";
        }
        if(getAllText(doc) != savedFileTxt) {
            return "Modified";
        } else {
            return "Saved";
        }
    } else {
        return "Not Saved";
    }
}

int copyy = -1, copyx = -1;

void moveCursor(int x, int y) {
    cursorX += x;
    cursorY += y;

    // Clamp cursor position to grid bounds
    cursorX = max(0, min(cursorX, numCols - 1));
    cursorY = max(0, min(cursorY, numRows - 1));

    return;
}

// Function to clear the input buffer
void clearInputBuffer() {
    while (_kbhit()) {
        _getch(); // Discard any pending key presses
    }
}

bool isKeyPressed(int key) {
    return (GetAsyncKeyState(key) & 0x8000) != 0;
}

// Custom getline function that doesn't create a new line
string getlineNoNewLine(int x, int y) {
    string input;
    char ch;
    DWORD mode;
    HANDLE hStdin = GetStdHandle(STD_INPUT_HANDLE);

    // Save the current console mode
    GetConsoleMode(hStdin, &mode);

    // Disable line input and echo input
    SetConsoleMode(hStdin, mode & ~(ENABLE_LINE_INPUT | ENABLE_ECHO_INPUT));

    // Read input character by character
    while (true) {
        ch = _getch(); // Read a single character
        if (ch == '\r') { // Enter key
            break;
        } else if (ch == '\b') { // Backspace key
            if (!input.empty()) {
                input.pop_back(); // Remove the last character
                gotoxy(x + input.length(), y); // Move cursor back
                cout << ' '; // Clear the character
                gotoxy(x + input.length(), y); // Move cursor back again
            }
        } else {
            input += ch; // Add the character to the input string
            cout << ch; // Display the character
        }
    }

    // Restore the original console mode
    SetConsoleMode(hStdin, mode);

    return input;
}

void stoplog(string message) {
    system("cls");
    cout << message << endl;
    system("pause");
}

void createTab(Document& document, const string& tabName) {
    // Crear una tabla vacía
    Table table(numRows, numCols); // Puedes inicializar la tabla con celdas vacías o con valores predeterminados

    // Crear una pestaña con la tabla y el nombre proporcionado
    Tab newTab(table, tabName);

    // Agregar la pestaña al documento
    document.getTabs().push_back(newTab);
}

void showprintconfirm(const vector<string>& lines, Document openedDocument) {
    system("cls");

    // Obtener el nombre del documento
    String name = openedDocument.getName();
    string title = " eLite NT Office 2025 Tablas - " + name + " | Confirm Print ";

    // Dibujar el título
    gotoxy(0, 0);
    string tbar = "";
    for (int i = 0; i < getConsoleResolution()[0] / 2 - title.length() / 2; i++) {
        tbar += " ";
    }
    cout << ASCII_DIM << ASCII_GREEN << ASCII_REVERSE << tbar << ASCII_RESET;
    cout << ASCII_GREEN << ASCII_REVERSE << title << ASCII_RESET;
    tbar = "";
    for (int i = 0; i < getConsoleResolution()[0] / 2 - title.length() / 2 - 1; i++) {
        tbar += " ";
    }
    cout << ASCII_DIM << ASCII_GREEN << ASCII_REVERSE << tbar << ASCII_RESET;
    cout << endl;

    for(string line : lines) {
        cout << line << endl;
    }

    switch (selectiondialog(getConsoleResolution()[0] / 2 - 50 / 2, getConsoleResolution()[1] / 2 - 6 / 2, 50, 6, "Print", {"Are you sure that you want to print this?"}, {"Yes","No","Cancel"}, rgbBackgrounds(240,240,240), ASCII_BLACK)) {
        case 0:
            print(lines);
        case 1:
            break;
        case 2:
            break;
    }
}

int main(int argc, char** argv) {
    system("cls");
    
    HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE);
    CONSOLE_SCREEN_BUFFER_INFO consoleInfo;
    GetConsoleScreenBufferInfo(hConsole, &consoleInfo);
    WORD originalAttributes = consoleInfo.wAttributes;

    Document openedDocument;

    if(argc == 1) {
        // Crear la pestaña inicial
        createInitialTab(openedDocument);

        openedDocument.name = "Untitled";

        // Verificar que la pestaña se haya creado
        if (openedDocument.getTabs().empty()) {
            std::cerr << "Error: No se pudo crear la pestaña inicial." << std::endl;
            return 1;
        }
    } else {
        string filenmae(argv[1]);

        filename = filenmae;

        if (!exists_test0(filename)) {
            openedDocument.name = filename;

            createInitialTab(openedDocument);
        } else {
            openedDocument = openDocument(filename, openedDocument);
        }
    }

    cout << ASCII_RESET;
    openedFolder = ".";

    // Draw the initial grid
    drawGrid(cursorX, cursorY, openedDocument);
    String ref;
    String savedFileTxt, l;

    bool exit = false;
    bool repaint = true;

    // Main program loop
    while (!exit) {
        for(int i = 0; i < getConsoleResolution()[0]; i++) {
            cout << ASCII_BG_WHITE << " " << ASCII_RESET;
        }
        gotoxy(0,getConsoleResolution()[1]-1);
        cout << ASCII_BLACK << ASCII_BG_WHITE << " " << getFileState(openedDocument) <<  ASCII_RESET;
        
        gotoxy(0,0);
        char input = _getch(); // Get user input without echoing it to the console

        // Handle user input
        string com;
        
        vector<int> consoleRes = getConsoleResolution();
        int consoleWidth = consoleRes[0];
        int consoleHeight = consoleRes[1];

        char nextcode;

        vector<string> menu;
        string menubar(30, ' ');
        int sel1 = 0;
        bool selected1 = false;

        switch (input) {
            case 75: // Move cursor left
                if (cursorX > 0) {
                    cursorX--;
                }
                break;
            case 77: // Move cursor right
                if (cursorX < numCols - 1) {
                    cursorX++;
                }
                break;
            case 72: // Move cursor up
                if (cursorY > 0) {
                    cursorY--;
                }
                break;
            case 80: // Move cursor down
                if (cursorY < numRows - 1) {
                    cursorY++;
                }
                break;
            case 'w': // Move cursor up
                if (cursorY > 0) {
                    cursorY--;
                }
                break;
            case 's': // Move cursor down
                if (cursorY < numRows - 1) {
                    cursorY++;
                }
                break;
            case 'a': // Move cursor left
                if (cursorX > 0) {
                    cursorX--;
                }
                break;
            case 'd': // Move cursor right
                if (cursorX < numCols - 1) {
                    cursorX++;
                }
                break;
            case '\t': // Move cursor to the next cell
                openedTab++;
                if (openedTab >= openedDocument.getTabs().size()) {
                    openedTab = 0;
                }
                break;
            case 0:
                nextcode = getch();

                switch(nextcode) {
                    case 60:
                        if (cursorX >= 0 && cursorX < numCols && cursorY >= 0 && cursorY < numRows) {
                            // Limpiar el contenido anterior de la celda
                            gotoxy(cursorX * 10 + 5, cursorY + 4);
                            cout << ASCII_UNDERLINE << "\033[48;2;16;93;201m" << "          ";

                            // Mover el cursor a la posición de la celda
                            gotoxy(cursorX * 10 + 5, cursorY + 4);

                            // Leer la nueva entrada del usuario
                            cout << ASCII_BLINK; // Activar el parpadeo
                            string newValue;
                            
                            getline(cin, newValue);

                            cout << ASCII_RESET; // Desactivar el parpadeo

                            // Actualizar el contenido de la celda en el documento
                            openedDocument.updateCell(openedTab, cursorY, cursorX, newValue);

                            drawGrid(cursorX, cursorY, openedDocument);
                        } else {
                            system("cls");
                            cout << "Error: Cursor fuera de límites." << endl;
                            system("pause");
                        }
                        break;
                    case 62:
                        if(getFileState(openedDocument) == "Modified" || getFileState(openedDocument) == "Not Saved") {        
                            int option = selectiondialog(getConsoleResolution()[0] / 2 - 50 / 2, getConsoleResolution()[1] / 2 - 6 / 2, 50, 6, "Exit", {"Do you want to save changes before leaving?"}, {"Yes","No","Cancel"}, rgbBackgrounds(240,240,240), ASCII_BLACK);
                            if(option == 0) {
                                saveDocument(openedDocument);
                                exit = true;
                            } else if(option == 1) {
                                exit = true;
                            }
                        } else {
                            exit = true;
                        }
                        break;
                    case 63:
                        system("cls");
                        drawGrid(cursorX, cursorY, openedDocument);
                        break;
                    case 68:
                        // Show menu
                        gotoxy(0,3);
                        menu = {"File", "Cell", "Sheet", "Print", "Help", "About"};
                        sel1 = 0;
                        
                        while(!selected1) {
                            gotoxy(0,3);
                            for(int i = 0; i < menu.size(); i++) {
                                cout << rgbBackgrounds(210,210,210) << menubar << ASCII_RESET << endl;
                            }
                            gotoxy(0,3);
                            for(int i = 0; i < menu.size(); i++) {
                                cout << rgbBackgrounds(210,210,210);
                                if(i == sel1) {
                                    cout << ASCII_BLACK << " " << menu[i] << " " << ASCII_RESET << endl;
                                } else {
                                    cout << " " << menu[i] << " " << endl;
                                }
                            }
                            
                            switch(_getch()) {
                                case 72:
                                    if(sel1 > 0) {
                                        sel1--;
                                    }
                                    break;
                                case 80:
                                    if(sel1 < menu.size() - 1) {
                                        sel1++;
                                    }
                                    break;
                                case 13:
                                    selected1 = true;
                                    break;
                                case 27:
                                    sel1 = -1;
                                    selected1 = true;
                                    break;
                            }
                        }

                        switch(sel1) {
                            case 0:
                                // File
                                menu = {"Open", "Save", "Save As", "Exit"};
                                sel1 = 0;
                                selected1 = false;

                                while(!selected1) {
                                    gotoxy(30,3);
                                    for(int i = 0; i < menu.size(); i++) {
                                        gotoxy(30, 3 + i);
                                        cout << rgbBackgrounds(210,210,210) << menubar << ASCII_RESET;
                                    }
                                    gotoxy(30,3);
                                    for(int i = 0; i < menu.size(); i++) {
                                        cout << rgbBackgrounds(210,210,210);
                                        gotoxy(30, 3 + i);
                                        if(i == sel1) {
                                            cout << ASCII_BLACK << " " << menu[i] << " " << ASCII_RESET;
                                        } else {
                                            cout << " " << menu[i] << " ";
                                        }
                                    }

                                    switch(_getch()) {
                                        case 72:
                                            if(sel1 > 0) {
                                                sel1--;
                                            }
                                            break;
                                        case 80:
                                            if(sel1 < menu.size() - 1) {
                                                sel1++;
                                            }
                                            break;
                                        case 13:
                                            selected1 = true;
                                            break;
                                        case 27:
                                            sel1 = -1;
                                            selected1 = true;
                                            break;
                                    }
                                }

                                switch(sel1) {
                                    case 0:
                                        // Open
                                        openedDocument = openDocument(inputDialog(getConsoleResolution()[0] / 2 - 50 / 2, getConsoleResolution()[1] / 2 - 7 / 2, 50, 7, "Open file", {"Enter filename to open: "}, rgbBackgrounds(240,240,240), ASCII_BLACK), openedDocument);
                                        break;
                                    case 1:
                                        // Save
                                        saveDocument(openedDocument);
                                        break;
                                    case 2:
                                        // Save As
                                        saveasDocument(openedDocument);
                                        break;
                                    case 3:
                                        // Exit
                                        if(getFileState(openedDocument) == "Modified" || getFileState(openedDocument) == "Not Saved") {        
                                            int option = selectiondialog(getConsoleResolution()[0] / 2 - 50 / 2, getConsoleResolution()[1] / 2 - 6 / 2, 50, 6, "Exit", {"Do you want to save changes before leaving?"}, {"Yes","No","Cancel"}, rgbBackgrounds(240,240,240), ASCII_BLACK);
                                            if(option == 0) {
                                                saveDocument(openedDocument);
                                                exit = true;
                                            } else if(option == 1) {
                                                exit = true;
                                            }
                                        } else {
                                            exit = true;
                                        }
                                        break;
                                }
                                break;
                            case 1:
                                // Cell
                                menu = {"Edit", "Copy", "Paste", "Delete"};
                                sel1 = 0;
                                selected1 = false;

                                while(!selected1) {
                                    gotoxy(30,4);
                                    for(int i = 0; i < menu.size(); i++) {
                                        gotoxy(30, 4 + i);
                                        cout << rgbBackgrounds(210,210,210) << menubar << ASCII_RESET;
                                    }
                                    gotoxy(30,4);
                                    for(int i = 0; i < menu.size(); i++) {
                                        cout << rgbBackgrounds(210,210,210);
                                        gotoxy(30, 4 + i);
                                        if(i == sel1) {
                                            cout << ASCII_BLACK << " " << menu[i] << " " << ASCII_RESET;
                                        } else {
                                            cout << " " << menu[i] << " ";
                                        }
                                    }

                                    switch(_getch()) {
                                        case 72:
                                            if(sel1 > 0) {
                                                sel1--;
                                            }
                                            break;
                                        case 80:
                                            if(sel1 < menu.size() - 1) {
                                                sel1++;
                                            }
                                            break;
                                        case 13:
                                            selected1 = true;
                                            break;
                                        case 27:
                                            sel1 = -1;
                                            selected1 = true;
                                            break;
                                    }
                                }

                                switch(sel1) {
                                    case 0:
                                        // Edit
                                        drawGrid(cursorX, cursorY, openedDocument);
                                        if (cursorX >= 0 && cursorX < numCols && cursorY >= 0 && cursorY < numRows) {
                                            // Limpiar el contenido anterior de la celda
                                            gotoxy(cursorX * 10 + 5, cursorY + 4);
                                            cout << ASCII_UNDERLINE << "\033[48;2;16;93;201m" << "          ";

                                            // Mover el cursor a la posición de la celda
                                            gotoxy(cursorX * 10 + 5, cursorY + 4);

                                            // Leer la nueva entrada del usuario
                                            cout << ASCII_BLINK; // Activar el parpadeo
                                            string newValue;
                                            
                                            getline(cin, newValue);

                                            cout << ASCII_RESET; // Desactivar el parpadeo

                                            // Actualizar el contenido de la celda en el documento
                                            openedDocument.updateCell(openedTab, cursorY, cursorX, newValue);

                                            drawGrid(cursorX, cursorY, openedDocument);
                                        } else {
                                            system("cls");
                                            cout << "Error: Cursor fuera de límites." << endl;
                                            system("pause");
                                        }
                                        break;
                                    case 1:
                                        // Copy
                                        copyy = cursorY;
                                        copyx = cursorX;
                                        break;
                                    case 2:
                                        // Paste
                                        if (copyy != -1 && copyx != -1) {
                                            // Copiar el contenido de la celda (copyy, copyx) a la celda actual (cursorY, cursorX)
                                            if (openedDocument.getTabs().size() > openedTab) {
                                                Table& table = openedDocument.getTabs()[openedTab].getTable();
                                                openedDocument.getTabs()[openedTab].getTable().getCell(cursorY, cursorX).setContent(table.getCell(copyy, copyx).getContent());
                                            }
                                        }
                                        break;
                                    case 3:
                                        // Delete
                                        openedDocument.updateCell(openedTab, cursorY, cursorX, "");
                                        drawGrid(cursorX, cursorY, openedDocument);
                                        break;
                                }
                                break;
                            case 2:
                                // Sheet
                                menu = {"Add", "Rename", "Delete"};
                                sel1 = 0;
                                selected1 = false;

                                while(!selected1) {
                                    gotoxy(30,5);
                                    for(int i = 0; i < menu.size(); i++) {
                                        gotoxy(30, 5 + i);
                                        cout << rgbBackgrounds(210,210,210) << menubar << ASCII_RESET;
                                    }
                                    gotoxy(30,5);
                                    for(int i = 0; i < menu.size(); i++) {
                                        cout << rgbBackgrounds(210,210,210);
                                        gotoxy(30, 5 + i);
                                        if(i == sel1) {
                                            cout << ASCII_BLACK << " " << menu[i] << " " << ASCII_RESET;
                                        } else {
                                            cout << " " << menu[i] << " ";
                                        }
                                    }

                                    switch(_getch()) {
                                        case 72:
                                            if(sel1 > 0) {
                                                sel1--;
                                            }
                                            break;
                                        case 80:
                                            if(sel1 < menu.size() - 1) {
                                                sel1++;
                                            }
                                            break;
                                        case 13:
                                            selected1 = true;
                                            break;
                                        case 27:
                                            sel1 = -1;
                                            selected1 = true;
                                            break;
                                    }
                                }

                                switch(sel1) {
                                    case 0:
                                        // Add
                                        openedDocument.getTabs().push_back(Tab(Table(numRows, numCols), inputDialog(getConsoleResolution()[0] / 2 - 50 / 2, getConsoleResolution()[1] / 2 - 7 / 2, 50, 7, "Add sheet", {"Enter sheet name: "}, rgbBackgrounds(240,240,240), ASCII_BLACK)));
                                        break;
                                    case 1:
                                        // Rename
                                        openedDocument.getTabs()[openedTab].name = inputDialog(getConsoleResolution()[0] / 2 - 50 / 2, getConsoleResolution()[1] / 2 - 7 / 2, 50, 7, "Rename sheet", {"Enter new sheet name: "}, rgbBackgrounds(240,240,240), ASCII_BLACK);
                                        break;
                                    case 2:
                                        // Delete
                                        if(openedDocument.getTabs().size() == 1) {
                                            errordialog(getConsoleResolution()[0] / 2 - 50 / 2, getConsoleResolution()[1] / 2 - 6 / 2, 50, 6, "Delete sheet", {"You can't delete the last sheet on the document."}, rgbBackgrounds(240,240,240), ASCII_BLACK);
                                            break;
                                        }

                                        if(selectiondialog(getConsoleResolution()[0] / 2 - 50 / 2, getConsoleResolution()[1] / 2 - 6 / 2, 50, 6, "Delete sheet", {"Are you sure you want to delete this sheet?"}, {"Yes","No"}, rgbBackgrounds(240,240,240), ASCII_BLACK) == 0) {
                                            openedDocument.getTabs().erase(openedDocument.getTabs().begin() + openedTab);
                                        }
                                        break;
                                }
                                break;
                            case 3:
                                // Print
                                switch(bigselectiondialog(getConsoleResolution()[0] / 2 - 50 / 2, getConsoleResolution()[1] / 2 - 8 / 2, 50, 8, "Print", {"Select what to print:"}, {"All the sheet","Whole document","Range of columns and rows of a specified sheet", "Cancel"}, rgbBackgrounds(240,240,240), ASCII_BLACK)) {
                                    case 0:
                                        // All the sheet

                                        showprintconfirm(openedDocument.getTabs()[openedTab].getTable().getPrinteableText(), openedDocument);
                                        break;
                                    case 1:
                                        // Whole document
                                        break;
                                    case 2:
                                        // Range of columns and rows
                                        break;
                                }
                                
                                break;
                            case 4:
                                // Help
                                drawGrid(cursorX, cursorY, openedDocument);
                                normaldialog(getConsoleResolution()[0] / 2 - 52 / 2, getConsoleResolution()[1] / 2 - 10 / 2, 52, 10, "Help", {"Welcome to eLite NT Office 2025 Tablas!", "At the moment, no online documentation is available", "Wait for the program be open-source to have", "documentation"}, rgbBackgrounds(240,240,240), ASCII_BLACK);
                                break;
                            case 5:
                                // About
                                normaldialog(getConsoleResolution()[0] / 2 - 50 / 2, getConsoleResolution()[1] / 2 - 10 / 2, 50, 10, "About", {"eLite NT Office 2025 Tablas", "(c) 2025 Blas Fernández", "v_2025.3.15"}, rgbBackgrounds(240,240,240), ASCII_BLACK);
                                break;
                        }
                        break;
                    case 69:
                        saveasDocument(openedDocument);
                        break;
                }
                break;
            case 'o': // Open file
                openedDocument = openDocument(inputDialog(getConsoleResolution()[0] / 2 - 50 / 2, getConsoleResolution()[1] / 2 - 7 / 2, 50, 7, "Open file", {"Enter filename to open: "}, rgbBackgrounds(240,240,240), ASCII_BLACK), openedDocument);
                break;
            case 'z': // Save file
                saveDocument(openedDocument);
                break;
            case 'x': // Save as
                saveasDocument(openedDocument);
                // Implement "Save As" functionality
                break;
            case '\r': // Enter key (Edit cell)
                {
                    // Verificar límites del cursor
                    if (cursorX >= 0 && cursorX < numCols && cursorY >= 0 && cursorY < numRows) {
                        // Limpiar el contenido anterior de la celda
                        gotoxy(cursorX * 10 + 5, cursorY + 4);
                        cout << ASCII_UNDERLINE << "\033[48;2;16;93;201m" << "              ";

                        // Mover el cursor a la posición de la celda
                        gotoxy(cursorX * 15 + 5, cursorY + 4);

                        // Leer la nueva entrada del usuario
                        cout << ASCII_BLINK; // Activar el parpadeo
                        string newValue;
                        
                        getline(cin, newValue);

                        cout << ASCII_RESET; // Desactivar el parpadeo

                        // Actualizar el contenido de la celda en el documento
                        openedDocument.updateCell(openedTab, cursorY, cursorX, newValue);

                        drawGrid(cursorX, cursorY, openedDocument);
                    } else {
                        system("cls");
                        cout << "Error: Cursor fuera de límites." << endl;
                        system("pause");
                    }
                }
                break;
            case 'q': // Exit
                if(getFileState(openedDocument) == "Modified" || getFileState(openedDocument) == "Not Saved") {        
                    int option = selectiondialog(getConsoleResolution()[0] / 2 - 50 / 2, getConsoleResolution()[1] / 2 - 6 / 2, 50, 6, "Exit", {"Do you want to save changes before leaving?"}, {"Yes","No","Cancel"}, rgbBackgrounds(240,240,240), ASCII_BLACK);
                    if(option == 0) {
                        saveDocument(openedDocument);
                        exit = true;
                    } else if(option == 1) {
                        exit = true;
                    }
                } else {
                    exit = true;
                }
                break;
            case ':':
                // Enter command mode
                gotoxy(getFileState(openedDocument).length() + 2, getConsoleResolution()[1]-1);
                cout << ASCII_BOLD << ASCII_BLACK << ASCII_BG_WHITE << ":";

                com = getlineNoNewLine(getFileState(openedDocument).length() + 3, consoleHeight - 1); // Start at column 18, row 0

                cout << ASCII_RESET;

                if (com == "save" || com == "g" || com == "s") {
                    saveDocument(openedDocument); // Guardar el documento actual
                } else if (com == "exit" || com == "q" || com == "e") {
                    if(getFileState(openedDocument) == "Modified" || getFileState(openedDocument) == "Not Saved") {        
                        int option = selectiondialog(getConsoleResolution()[0] / 2 - 50 / 2, getConsoleResolution()[1] / 2 - 6 / 2, 50, 6, "Exit", {"Do you want to save changes before leaving?"}, {"Yes","No","Cancel"}, rgbBackgrounds(240,240,240), ASCII_BLACK);
                        if(option == 0) {
                            saveDocument(openedDocument);
                            exit = true;
                        } else if(option == 1) {
                            exit = true;
                        }
                    } else {
                        exit = true;
                    }
                } else if (com == "copy" || com == "c") {
                    copyy = cursorY;
                    copyx = cursorX;
                } else if (com == "paste" || com == "p") {
                    if (copyy != -1 && copyx != -1) {
                        // Copiar el contenido de la celda (copyy, copyx) a la celda actual (cursorY, cursorX)
                        if (openedDocument.getTabs().size() > openedTab) {
                            Table& table = openedDocument.getTabs()[openedTab].getTable();
                            openedDocument.getTabs()[openedTab].getTable().getCell(cursorY, cursorX).setContent(table.getCell(copyy, copyx).getContent());
                        }
                    }
                } else if (com == "saveas" || com == "sa") {
                    saveasDocument(openedDocument); // Guardar el documento actual como un nuevo archivo
                } else if (com == "open" || com == "o") {
                    openedDocument = openDocument(inputDialog(getConsoleResolution()[0] / 2 - 50 / 2, getConsoleResolution()[1] / 2 - 6 / 2, 50, 6, "Open file", {"Enter filename to open: "}, rgbBackgrounds(240,240,240), ASCII_BLACK), openedDocument); // Abrir un archivo
                } else if (com == "help" || com == "h") {
                    system("cls");
                    cout << "Commands:\n";
                    cout << "save (s, g) - Save the current file\n";
                    cout << "saveas (sa) - Save the current file as a new file\n";
                    cout << "open (o) - Open a file\n";
                    cout << "exit (e, q) - Exit the program\n";
                    cout << "copy (c) - Copy the selected cell\n";
                    cout << "paste (p) - Paste the copied cell\n";
                    cout << "help (h) - Show this help message\n";
                    cout << "about - Show information about the program\n";
                    cout << "Press any key to continue...";
                    getch();
                } else if(com == "about") {
                    normaldialog(getConsoleResolution()[0] / 2 - 50 / 2, getConsoleResolution()[1] / 2 - 10 / 2, 50, 10, "About", {"eLite NT Office 2025 Tablas", "(c) 2025 Blas Fernández", "v_2025.3.15"}, rgbBackgrounds(240,240,240), ASCII_BLACK);
                }
                break;
            default:
                repaint = false;
                break;
        }

        // Redraw the grid after each input
        try {
            if(repaint) {
                drawGrid(cursorX, cursorY, openedDocument);
            }
            repaint = true;
        } catch(bad_alloc) {
            system("CLS");
            cursorX = 0;
            cursorY = 0;
            drawGrid(0, 0, openedDocument);
        }
    }

    system("cls");

    return 0;
}

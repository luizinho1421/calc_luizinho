// Reconstructed C++ Source Code of Windows 10/11 calc.exe (10.0.19041.1)
// Disassembled from RVA 0x113C (wWinMain)

#include <windows.h>
#include <shellapi.h>
#include <TraceLoggingProvider.h>

// Provider: "MicrosoftCalculator"
// GUID: {B6401E61-212F-8029-B9E0-29004D696372} (from .rdata 0x1400023C4)
TRACELOGGING_DEFINE_PROVIDER(
    g_hCalculatorProvider,
    "MicrosoftCalculator",
    (0xb6401e61, 0x212f, 0x8029, 0xb9, 0xe0, 0x29, 0x00, 0x4d, 0x69, 0x63, 0x72));

int WINAPI wWinMain(
    _In_ HINSTANCE hInstance,
    _In_opt_ HINSTANCE hPrevInstance,
    _In_ LPWSTR lpCmdLine,
    _In_ int nShowCmd)
{
    TraceLoggingRegister(g_hCalculatorProvider);

    TraceLoggingWrite(
        g_hCalculatorProvider,
        "CalculatorWinMain",
        TraceLoggingValue("CalculatorStarted", "CalculatorStarted"),
        TraceLoggingKeyword(0x200000000000));

    ShellExecuteW(NULL, NULL, L"ms-calculator:", NULL, NULL, SW_SHOWNORMAL);

    return 0;
}

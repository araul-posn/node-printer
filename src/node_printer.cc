#include "node_printer.hpp"

Napi::Object Init(Napi::Env env, Napi::Object exports) {
    exports.Set(Napi::String::New(env, "getPrinters"), Napi::Function::New(env, getPrinters));
    exports.Set(Napi::String::New(env, "getDefaultPrinterName"), Napi::Function::New(env, getDefaultPrinterName));
    exports.Set(Napi::String::New(env, "getPrinter"), Napi::Function::New(env, getPrinter));
    exports.Set(Napi::String::New(env, "getPrinterDriverOptions"), Napi::Function::New(env, getPrinterDriverOptions));
    exports.Set(Napi::String::New(env, "getJob"), Napi::Function::New(env, getJob));
    exports.Set(Napi::String::New(env, "setJob"), Napi::Function::New(env, setJob));
    exports.Set(Napi::String::New(env, "printDirect"), Napi::Function::New(env, PrintDirect));
    exports.Set(Napi::String::New(env, "printFile"), Napi::Function::New(env, PrintFile));
    exports.Set(Napi::String::New(env, "getSupportedPrintFormats"), Napi::Function::New(env, getSupportedPrintFormats));
    exports.Set(Napi::String::New(env, "getSupportedJobCommands"), Napi::Function::New(env, getSupportedJobCommands));
    
    return exports;
}

NODE_API_MODULE(node_printer, Init)

// Helpers

bool getStringOrBufferFromV8Value(Napi::Value iV8Value, std::string &oData)
{
    if(iV8Value.IsString())
    {
        oData = iV8Value.As<Napi::String>().Utf8Value();
        return true;
    }
    if(iV8Value.IsBuffer())
    {
        Napi::Buffer<char> buffer = iV8Value.As<Napi::Buffer<char>>();
        oData.assign(buffer.Data(), buffer.Length());
        return true;
    }
    return false;
}
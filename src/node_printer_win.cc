#include "node_printer.hpp"

#if _MSC_VER
#include <windows.h>
#include <Winspool.h>
#include <Wingdi.h>
#pragma  comment(lib, "Winspool.lib")
#else
#error "Unsupported compiler for windows. Feel free to add it."
#endif

#include <string>
#include <map>
#include <utility>
#include <sstream>

namespace{
    typedef std::map<std::string, DWORD> StatusMapType;

    /** Memory value class management to avoid memory leak
    */
    template<typename Type>
    class MemValue: public MemValueBase<Type> {
    public:
        /** Constructor of allocating iSizeKbytes bytes memory;
        * @param iSizeKbytes size in bytes of required allocating memory
        */
        MemValue(const DWORD iSizeKbytes) {
            this->_value = (Type*)malloc(iSizeKbytes);
        }
		
        ~MemValue () {
            free();
        }
    protected:
        virtual void free() {
            if(this->_value != NULL)
            {
                ::free(this->_value);
                this->_value = NULL;
            }
        }
    };

    struct PrinterHandle
    {
        PrinterHandle(LPWSTR iPrinterName)
        {
            _ok = OpenPrinterW(iPrinterName, &_printer, NULL);
        }
        ~PrinterHandle()
        {
            if(_ok)
            {
                ClosePrinter(_printer);
            }
        }
        operator HANDLE() {return _printer;}
        operator bool() { return (!!_ok);}
        HANDLE & operator *() { return _printer;}
        HANDLE * operator ->() { return &_printer;}
        const HANDLE & operator ->() const { return _printer;}
        HANDLE _printer;
        BOOL _ok;
    };

    const StatusMapType& getStatusMap()
    {
        static StatusMapType result;
        if(!result.empty())
        {
            return result;
        }
        // add only first time
#define STATUS_PRINTER_ADD(value, type) result.insert(std::make_pair(value, type))
        STATUS_PRINTER_ADD("BUSY", PRINTER_STATUS_BUSY);
        STATUS_PRINTER_ADD("DOOR-OPEN", PRINTER_STATUS_DOOR_OPEN);
        STATUS_PRINTER_ADD("ERROR", PRINTER_STATUS_ERROR);
        STATUS_PRINTER_ADD("INITIALIZING", PRINTER_STATUS_INITIALIZING);
        STATUS_PRINTER_ADD("IO-ACTIVE", PRINTER_STATUS_IO_ACTIVE);
        STATUS_PRINTER_ADD("MANUAL-FEED", PRINTER_STATUS_MANUAL_FEED);
        STATUS_PRINTER_ADD("NO-TONER", PRINTER_STATUS_NO_TONER);
        STATUS_PRINTER_ADD("NOT-AVAILABLE", PRINTER_STATUS_NOT_AVAILABLE);
        STATUS_PRINTER_ADD("OFFLINE", PRINTER_STATUS_OFFLINE);
        STATUS_PRINTER_ADD("OUT-OF-MEMORY", PRINTER_STATUS_OUT_OF_MEMORY);
        STATUS_PRINTER_ADD("OUTPUT-BIN-FULL", PRINTER_STATUS_OUTPUT_BIN_FULL);
        STATUS_PRINTER_ADD("PAGE-PUNT", PRINTER_STATUS_PAGE_PUNT);
        STATUS_PRINTER_ADD("PAPER-JAM", PRINTER_STATUS_PAPER_JAM);
        STATUS_PRINTER_ADD("PAPER-OUT", PRINTER_STATUS_PAPER_OUT);
        STATUS_PRINTER_ADD("PAPER-PROBLEM", PRINTER_STATUS_PAPER_PROBLEM);
        STATUS_PRINTER_ADD("PAUSED", PRINTER_STATUS_PAUSED);
        STATUS_PRINTER_ADD("PENDING-DELETION", PRINTER_STATUS_PENDING_DELETION);
        STATUS_PRINTER_ADD("POWER-SAVE", PRINTER_STATUS_POWER_SAVE);
        STATUS_PRINTER_ADD("PRINTING", PRINTER_STATUS_PRINTING);
        STATUS_PRINTER_ADD("PROCESSING", PRINTER_STATUS_PROCESSING);
        STATUS_PRINTER_ADD("SERVER-UNKNOWN", PRINTER_STATUS_SERVER_UNKNOWN);
        STATUS_PRINTER_ADD("TONER-LOW", PRINTER_STATUS_TONER_LOW);
        STATUS_PRINTER_ADD("USER-INTERVENTION", PRINTER_STATUS_USER_INTERVENTION);
        STATUS_PRINTER_ADD("WAITING", PRINTER_STATUS_WAITING);
        STATUS_PRINTER_ADD("WARMING-UP", PRINTER_STATUS_WARMING_UP);
        //JOB status
        STATUS_PRINTER_ADD("PRINTING", JOB_STATUS_PRINTING);
        STATUS_PRINTER_ADD("PAUSED", JOB_STATUS_PAUSED);
        STATUS_PRINTER_ADD("ERROR", JOB_STATUS_ERROR);
        STATUS_PRINTER_ADD("PENDING-DELETION", JOB_STATUS_DELETING);
        STATUS_PRINTER_ADD("OFFLINE", JOB_STATUS_OFFLINE);
        STATUS_PRINTER_ADD("PAPEROUT", JOB_STATUS_PAPEROUT);
        STATUS_PRINTER_ADD("PRINTED", JOB_STATUS_PRINTED);
        STATUS_PRINTER_ADD("DELETED", JOB_STATUS_DELETED);
        STATUS_PRINTER_ADD("BLOCKED", JOB_STATUS_BLOCKED_DEVQ);
        STATUS_PRINTER_ADD("USER-INTERVENTION", JOB_STATUS_USER_INTERVENTION);
        // Custom
        STATUS_PRINTER_ADD("PENDING", JOB_STATUS_RESTART);
        STATUS_PRINTER_ADD("PENDING", JOB_STATUS_SPOOLING);
#undef STATUS_PRINTER_ADD
        return result;
    }

    const StatusMapType& getJobCommandMap()
    {
        static StatusMapType result;
        if(!result.empty())
        {
            return result;
        }
        // add only first time
#define STATUS_PRINTER_ADD(value, type) result.insert(std::make_pair(value, type))
        STATUS_PRINTER_ADD("CANCEL", JOB_CONTROL_CANCEL);
        STATUS_PRINTER_ADD("PAUSE", JOB_CONTROL_PAUSE);
        STATUS_PRINTER_ADD("RESTART", JOB_CONTROL_RESTART);
        STATUS_PRINTER_ADD("RESUME", JOB_CONTROL_RESUME);
        STATUS_PRINTER_ADD("DELETE", JOB_CONTROL_DELETE);
        STATUS_PRINTER_ADD("SENT-TO-PRINTER", JOB_CONTROL_SENT_TO_PRINTER);
        STATUS_PRINTER_ADD("LAST-PAGE-EJECTED", JOB_CONTROL_LAST_PAGE_EJECTED);
        STATUS_PRINTER_ADD("RETAIN", JOB_CONTROL_RETAIN);
        STATUS_PRINTER_ADD("RELEASE", JOB_CONTROL_RELEASE);
#undef STATUS_PRINTER_ADD
        return result;
    }

    /** Return the status as array of string
    */
    Napi::Array parseStatusArray(const DWORD iStatus, Napi::Env env)
    {
        Napi::Array result = Napi::Array::New(env);
        int i = 0;
        for(StatusMapType::const_iterator itStatus = getStatusMap().begin(); itStatus != getStatusMap().end(); ++itStatus)
        {
            if(itStatus->second & iStatus)
            {
                result.Set(i++, Napi::String::New(env, itStatus->first));
            }
        }
        return result;
    }

    /** Convert a wstring to a UTF8 string
    */
    std::string ConvertToUtf8String(const std::wstring &wstr)
    {
        int cbMultiByte = WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, NULL, 0, NULL, NULL);
        LPSTR lpMultiByteStr = (LPSTR)malloc(cbMultiByte);
        cbMultiByte = WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, lpMultiByteStr, cbMultiByte, NULL, NULL);
        std::string str(lpMultiByteStr);
        free(lpMultiByteStr);
        return str;
    }

    /** Convert a UTF8 string to a wstring
    */
    std::wstring ConvertToWString(const std::string &str)
    {
        int cchWideChar = MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, NULL, 0);
        LPWSTR lpWideCharStr = (LPWSTR)malloc(cchWideChar * sizeof(WCHAR));
        cchWideChar = MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, lpWideCharStr, cchWideChar);
        std::wstring wstr(lpWideCharStr);
        free(lpWideCharStr);
        return wstr;
    }

    void parsePrinterInfo(const PRINTER_INFO_2W *printer, Napi::Object& result, Napi::Env env)
    {
        result.Set("name", Napi::String::New(env, ConvertToUtf8String(printer->pPrinterName)));
        result.Set("serverName", (printer->pServerName == NULL) ? env.Undefined() : Napi::String::New(env, ConvertToUtf8String(printer->pServerName)));
        result.Set("shareName", (printer->pShareName == NULL) ? env.Undefined() : Napi::String::New(env, ConvertToUtf8String(printer->pShareName)));
        result.Set("portName", Napi::String::New(env, ConvertToUtf8String(printer->pPortName)));
        result.Set("driverName", Napi::String::New(env, ConvertToUtf8String(printer->pDriverName)));
        result.Set("comment", (printer->pComment == NULL) ? env.Undefined() : Napi::String::New(env, ConvertToUtf8String(printer->pComment)));
        result.Set("location", (printer->pLocation == NULL) ? env.Undefined() : Napi::String::New(env, ConvertToUtf8String(printer->pLocation)));
        result.Set("sepFile", (printer->pSepFile == NULL) ? env.Undefined() : Napi::String::New(env, ConvertToUtf8String(printer->pSepFile)));
        result.Set("printProcessor", Napi::String::New(env, ConvertToUtf8String(printer->pPrintProcessor)));
        result.Set("datatype", Napi::String::New(env, ConvertToUtf8String(printer->pDatatype)));
        result.Set("parameters", (printer->pParameters == NULL) ? env.Undefined() : Napi::String::New(env, ConvertToUtf8String(printer->pParameters)));
        result.Set("attributes", Napi::Number::New(env, printer->Attributes));
        result.Set("priority", Napi::Number::New(env, printer->Priority));
        result.Set("defaultPriority", Napi::Number::New(env, printer->DefaultPriority));
        result.Set("startTime", Napi::Number::New(env, printer->StartTime));
        result.Set("untilTime", Napi::Number::New(env, printer->UntilTime));
        result.Set("jobs", Napi::Number::New(env, printer->cJobs));
        result.Set("averagePpm", Napi::Number::New(env, printer->AveragePPM));
        result.Set("status", parseStatusArray(printer->Status, env));
        result.Set("statusNumber", Napi::Number::New(env, printer->Status));
        result.Set("isDefault", Napi::Boolean::New(env, false));
        Napi::Object options = Napi::Object::New(env);
        options.Set("copies", Napi::String::New(env, "1"));
        options.Set("media", Napi::String::New(env, "A4"));
        options.Set("n-up", Napi::String::New(env, "1"));
        result.Set("options", options);
    }

    void parseJobInfo(const JOB_INFO_2W *job, Napi::Object& result, Napi::Env env)
    {
        //Common fields
        result.Set("id", Napi::Number::New(env, job->JobId));
        result.Set("name", Napi::String::New(env, ConvertToUtf8String(job->pDocument)));
        result.Set("printerName", Napi::String::New(env, ConvertToUtf8String(job->pPrinterName)));
        result.Set("user", Napi::String::New(env, ConvertToUtf8String(job->pUserName)));
        result.Set("format", Napi::String::New(env, ConvertToUtf8String(job->pDatatype)));
        result.Set("priority", Napi::Number::New(env, job->Priority));
        result.Set("size", Napi::Number::New(env, job->Size));
        result.Set("status", parseStatusArray(job->Status, env));
        result.Set("statusNumber", Napi::Number::New(env, job->Status));
        
        //Specific fields
        SYSTEMTIME stJobTime;
        FILETIME ftJobTime;
        // XXX: Hack by casting DWORD to LARGE_INTEGER
        // http://stackoverflow.com/questions/353289/access-filetime-as-a-large-integer-dword-union-with-c
        reinterpret_cast<LARGE_INTEGER*>(&ftJobTime)->QuadPart = reinterpret_cast<const LARGE_INTEGER*>(&job->Submitted)->QuadPart;
        if(FileTimeToSystemTime(&ftJobTime, &stJobTime))
        {
            Napi::Object jsTime = Napi::Object::New(env);
            jsTime.Set("year", Napi::Number::New(env, stJobTime.wYear));
            jsTime.Set("month", Napi::Number::New(env, stJobTime.wMonth));
            jsTime.Set("dayOfWeek", Napi::Number::New(env, stJobTime.wDayOfWeek));
            jsTime.Set("day", Napi::Number::New(env, stJobTime.wDay));
            jsTime.Set("hour", Napi::Number::New(env, stJobTime.wHour));
            jsTime.Set("minute", Napi::Number::New(env, stJobTime.wMinute));
            jsTime.Set("second", Napi::Number::New(env, stJobTime.wSecond));
            jsTime.Set("milliseconds", Napi::Number::New(env, stJobTime.wMilliseconds));
            result.Set("time", jsTime);
        }
        result.Set("pagesPrinted", Napi::Number::New(env, job->PagesPrinted));
    }
}

Napi::Value getPrinters(const Napi::CallbackInfo& info)
{
    Napi::Env env = info.Env();
    
    //http://msdn.microsoft.com/en-us/library/windows/desktop/ms678405(v=vs.85).aspx
    DWORD printers_size = 0;
    DWORD printers_size_bytes = 0;
    DWORD Level = 2;
    DWORD flags = PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS; // https://msdn.microsoft.com/en-us/library/cc244669.aspx
    
    // First try to retrieve the number of printers
    DWORD result = EnumPrintersW(flags, NULL, Level, NULL, 0, &printers_size_bytes, &printers_size);
    // allocate the required memory
    MemValue<PRINTER_INFO_2W> printers(printers_size_bytes);
    if(!printers)
    {
        Napi::Error::New(env, "Error on allocating memory for printers").ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    // Receive the enum printers
    result = EnumPrintersW(flags, NULL, Level, (LPBYTE)(printers.get()), printers_size_bytes, &printers_size_bytes, &printers_size);
    if(!result)
    {
        std::string error_str("Error on EnumPrinters: code: ");
        std::ostringstream s;
        s << GetLastError();
        error_str += s.str();
        Napi::Error::New(env, error_str).ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    // Get default printer name
    WCHAR szDefaultPrinterName[MAX_PATH];
    DWORD cchDefaultPrinterName = MAX_PATH;
    GetDefaultPrinterW(szDefaultPrinterName, &cchDefaultPrinterName);
    
    Napi::Array result_array = Napi::Array::New(env);
    PRINTER_INFO_2W *printer = printers.get();
    DWORD i = 0;
    for(; i < printers_size; ++i, ++printer)
    {
        Napi::Object result_printer = Napi::Object::New(env);
        parsePrinterInfo(printer, result_printer, env);
        
        // Check if this is the default printer
        if(wcscmp(printer->pPrinterName, szDefaultPrinterName) == 0)
        {
            result_printer.Set("isDefault", Napi::Boolean::New(env, true));
        }
        
        PrinterHandle printerHandle((LPWSTR)printer->pPrinterName);
        if(!printerHandle)
        {
            result_printer.Set("jobs", Napi::Array::New(env));
            result_array.Set(i, result_printer);
            continue;
        }
        
        // Retrieve the number of jobs
        DWORD jobs_size = 0;
        DWORD jobs_size_bytes = 0;
        result = EnumJobsW(*printerHandle, 0, 0xFFFFFFFF, 2, NULL, 0, &jobs_size_bytes, &jobs_size);
        MemValue<JOB_INFO_2W> jobs(jobs_size_bytes);
        if(!jobs)
        {
            result_printer.Set("jobs", Napi::Array::New(env));
            result_array.Set(i, result_printer);
            continue;
        }
        result = EnumJobsW(*printerHandle, 0, 0xFFFFFFFF, 2, (LPBYTE)(jobs.get()), jobs_size_bytes, &jobs_size_bytes, &jobs_size);
        if(!result)
        {
            result_printer.Set("jobs", Napi::Array::New(env));
            result_array.Set(i, result_printer);
            continue;
        }
        
        Napi::Array result_jobs = Napi::Array::New(env);
        JOB_INFO_2W *job = jobs.get();
        for(DWORD j = 0; j < jobs_size; ++j, ++job)
        {
            Napi::Object result_job = Napi::Object::New(env);
            parseJobInfo(job, result_job, env);
            result_jobs.Set(j, result_job);
        }
        result_printer.Set("jobs", result_jobs);
        result_array.Set(i, result_printer);
    }
    
    return result_array;
}

Napi::Value getDefaultPrinterName(const Napi::CallbackInfo& info) 
{
    Napi::Env env = info.Env();
    WCHAR szDefaultPrinterName[MAX_PATH];
    DWORD cchDefaultPrinterName = MAX_PATH;
    GetDefaultPrinterW(szDefaultPrinterName, &cchDefaultPrinterName);
    return Napi::String::New(env, ConvertToUtf8String(szDefaultPrinterName));
}

Napi::Value getPrinter(const Napi::CallbackInfo& info)
{
    Napi::Env env = info.Env();
    
    if(info.Length() < 1)
    {
        Napi::TypeError::New(env, "getPrinter:invalid number of arguments (1 expected)").ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    if(!info[0].IsString())
    {
        Napi::TypeError::New(env, "getPrinter:first argument must be a string").ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    std::string printer_str = info[0].As<Napi::String>().Utf8Value();
    std::wstring printer_name = ConvertToWString(printer_str);
    
    PrinterHandle printerHandle((LPWSTR)printer_name.c_str());
    if(!printerHandle)
    {
        std::string error_str("Error on opening printer: ");
        error_str += printer_str;
        Napi::Error::New(env, error_str).ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    DWORD printer_info_size = 0;
    GetPrinterW(*printerHandle, 2, NULL, 0, &printer_info_size);
    if(!printer_info_size)
    {
        Napi::Error::New(env, "Error on GetPrinter to get printer info buffer size").ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    MemValue<PRINTER_INFO_2W> printer_info(printer_info_size);
    DWORD result = GetPrinterW(*printerHandle, 2, (LPBYTE)(printer_info.get()), printer_info_size, &printer_info_size);
    if(!result)
    {
        Napi::Error::New(env, "Error on GetPrinter").ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    Napi::Object result_printer = Napi::Object::New(env);
    parsePrinterInfo(printer_info.get(), result_printer, env);
    
    // Retrieve the number of jobs
    DWORD jobs_size = 0;
    DWORD jobs_size_bytes = 0;
    result = EnumJobsW(*printerHandle, 0, 0xFFFFFFFF, 2, NULL, 0, &jobs_size_bytes, &jobs_size);
    if(jobs_size > 0)
    {
        MemValue<JOB_INFO_2W> jobs(jobs_size_bytes);
        if(jobs)
        {
            result = EnumJobsW(*printerHandle, 0, 0xFFFFFFFF, 2, (LPBYTE)(jobs.get()), jobs_size_bytes, &jobs_size_bytes, &jobs_size);
            if(result)
            {
                Napi::Array result_jobs = Napi::Array::New(env);
                JOB_INFO_2W *job = jobs.get();
                for(DWORD j = 0; j < jobs_size; ++j, ++job)
                {
                    Napi::Object result_job = Napi::Object::New(env);
                    parseJobInfo(job, result_job, env);
                    result_jobs.Set(j, result_job);
                }
                result_printer.Set("jobs", result_jobs);
            }
        }
    }
    
    if(!result_printer.Has("jobs"))
    {
        result_printer.Set("jobs", Napi::Array::New(env));
    }
    
    return result_printer;
}

Napi::Value getPrinterDriverOptions(const Napi::CallbackInfo& info) 
{
    Napi::Env env = info.Env();
    
    if(info.Length() < 1)
    {
        Napi::TypeError::New(env, "getPrinterDriverOptions:invalid number of arguments (1 expected)").ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    if(!info[0].IsString())
    {
        Napi::TypeError::New(env, "getPrinterDriverOptions:first argument must be a string").ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    return Napi::Object::New(env);
}

Napi::Value getJob(const Napi::CallbackInfo& info) 
{
    Napi::Env env = info.Env();
    
    if(info.Length() < 2)
    {
        Napi::TypeError::New(env, "getJob:invalid number of arguments (2 expected)").ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    if(!info[0].IsString())
    {
        Napi::TypeError::New(env, "getJob:first argument must be a string").ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    if(!info[1].IsNumber())
    {
        Napi::TypeError::New(env, "getJob:second argument must be a number").ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    std::string printer_str = info[0].As<Napi::String>().Utf8Value();
    int job_id = info[1].As<Napi::Number>().Int32Value();
    std::wstring printer_name = ConvertToWString(printer_str);
    
    PrinterHandle printerHandle((LPWSTR)printer_name.c_str());
    if(!printerHandle)
    {
        std::string error_str("Error on opening printer: ");
        error_str += printer_str;
        Napi::Error::New(env, error_str).ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    DWORD job_info_size = 0;
    GetJobW(*printerHandle, job_id, 2, NULL, 0, &job_info_size);
    if(!job_info_size)
    {
        std::string error_str("Error on GetJob to get job info buffer size: ");
        std::ostringstream s;
        s << GetLastError();
        error_str += s.str();
        Napi::Error::New(env, error_str).ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    MemValue<JOB_INFO_2W> job_info(job_info_size);
    DWORD result = GetJobW(*printerHandle, job_id, 2, (LPBYTE)(job_info.get()), job_info_size, &job_info_size);
    if(!result)
    {
        Napi::Error::New(env, "Error on GetJob").ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    Napi::Object result_job = Napi::Object::New(env);
    parseJobInfo(job_info.get(), result_job, env);
    
    return result_job;
}

Napi::Value setJob(const Napi::CallbackInfo& info) 
{
    Napi::Env env = info.Env();
    
    if(info.Length() < 3)
    {
        Napi::TypeError::New(env, "setJob:invalid number of arguments (3 expected)").ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    if(!info[0].IsString())
    {
        Napi::TypeError::New(env, "setJob:first argument must be a string").ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    if(!info[1].IsNumber())
    {
        Napi::TypeError::New(env, "setJob:second argument must be a number").ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    if(!info[2].IsString())
    {
        Napi::TypeError::New(env, "setJob:third argument must be a string").ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    std::string printer_str = info[0].As<Napi::String>().Utf8Value();
    int job_id = info[1].As<Napi::Number>().Int32Value();
    std::string job_command_str = info[2].As<Napi::String>().Utf8Value();
    std::wstring printer_name = ConvertToWString(printer_str);
    
    PrinterHandle printerHandle((LPWSTR)printer_name.c_str());
    if(!printerHandle)
    {
        std::string error_str("Error on opening printer: ");
        error_str += printer_str;
        Napi::Error::New(env, error_str).ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    StatusMapType::const_iterator itJobCommand = getJobCommandMap().find(job_command_str);
    if(itJobCommand == getJobCommandMap().end())
    {
        Napi::Error::New(env, "wrong job command. use getSupportedJobCommands to see the possible commands").ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    DWORD result = SetJobW(*printerHandle, job_id, 0, NULL, itJobCommand->second);
    return Napi::Boolean::New(env, (result != 0));
}

Napi::Value PrintDirect(const Napi::CallbackInfo& info) 
{
    Napi::Env env = info.Env();
    
    if(info.Length() < 1)
    {
        Napi::TypeError::New(env, "printDirect:invalid number of arguments (1 expected)").ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    if(!info[0].IsObject())
    {
        Napi::TypeError::New(env, "printDirect:first argument must be an object").ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    Napi::Object arg_params = info[0].As<Napi::Object>();
    
    // check data property
    if(!arg_params.Has("data"))
    {
        Napi::TypeError::New(env, "printDirect:data parameter is mandatory").ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    // data
    Napi::Value arg_value_data = arg_params.Get("data");
    std::string data;
    if(!getStringOrBufferFromV8Value(arg_value_data, data))
    {
        Napi::TypeError::New(env, "printDirect:data parameter must be a string or Buffer").ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    std::wstring printer_name;
    // printer name
    if(arg_params.Has("printer"))
    {
        Napi::Value arg_value_printer = arg_params.Get("printer");
        if(!arg_value_printer.IsString())
        {
            Napi::TypeError::New(env, "printDirect:printer parameter must be a string").ThrowAsJavaScriptException();
            return env.Undefined();
        }
        std::string printer_str = arg_value_printer.As<Napi::String>().Utf8Value();
        printer_name = ConvertToWString(printer_str);
    }
    else
    {
        // if printer is not specified, then use default printer.
        WCHAR szDefaultPrinterName[MAX_PATH];
        DWORD cchDefaultPrinterName = MAX_PATH;
        GetDefaultPrinterW(szDefaultPrinterName, &cchDefaultPrinterName);
        printer_name = szDefaultPrinterName;
    }
    
    // type
    std::wstring docformat = L"RAW";
    if(arg_params.Has("type"))
    {
        Napi::Value arg_value_type = arg_params.Get("type");
        if(!arg_value_type.IsString())
        {
            Napi::TypeError::New(env, "printDirect:type parameter must be a string").ThrowAsJavaScriptException();
            return env.Undefined();
        }
        std::string type_str = arg_value_type.As<Napi::String>().Utf8Value();
        docformat = ConvertToWString(type_str);
    }
    
    // docname
    std::wstring docname = L"node print job";
    if(arg_params.Has("docname"))
    {
        Napi::Value arg_value_docname = arg_params.Get("docname");
        if(!arg_value_docname.IsString())
        {
            Napi::TypeError::New(env, "printDirect:docname parameter must be a string").ThrowAsJavaScriptException();
            return env.Undefined();
        }
        std::string docname_str = arg_value_docname.As<Napi::String>().Utf8Value();
        docname = ConvertToWString(docname_str);
    }
    
    // Open the printer
    PrinterHandle printerHandle((LPWSTR)printer_name.c_str());
    if(!printerHandle)
    {
        std::string error_str("Error on opening printer: ");
        error_str += ConvertToUtf8String(printer_name);
        Napi::Error::New(env, error_str).ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    // Start document
    DOC_INFO_1W docInfo;
    docInfo.pDocName = (LPWSTR)docname.c_str();
    docInfo.pOutputFile = NULL;
    docInfo.pDatatype = (LPWSTR)docformat.c_str();
    
    DWORD job_id = StartDocPrinterW(*printerHandle, 1, (LPBYTE)&docInfo);
    if(job_id == 0)
    {
        std::string error_str("Error on StartDocPrinter: ");
        std::ostringstream s;
        s << GetLastError();
        error_str += s.str();
        Napi::Error::New(env, error_str).ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    // Start page
    DWORD result = StartPagePrinter(*printerHandle);
    if(!result)
    {
        EndDocPrinter(*printerHandle);
        Napi::Error::New(env, "Error on StartPagePrinter").ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    // Write data
    DWORD bytes_written = 0;
    result = WritePrinter(*printerHandle, (LPVOID)data.c_str(), data.size(), &bytes_written);
    if(!result || bytes_written != data.size())
    {
        EndPagePrinter(*printerHandle);
        EndDocPrinter(*printerHandle);
        Napi::Error::New(env, "Error on WritePrinter").ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    // End page and document
    EndPagePrinter(*printerHandle);
    EndDocPrinter(*printerHandle);
    
    return Napi::Boolean::New(env, true);
}

Napi::Value PrintFile(const Napi::CallbackInfo& info) 
{
    Napi::Env env = info.Env();
    Napi::TypeError::New(env, "printFile() is not implemented yet on Windows.").ThrowAsJavaScriptException();
    return env.Undefined();
}

Napi::Value getSupportedPrintFormats(const Napi::CallbackInfo& info) 
{
    Napi::Env env = info.Env();
    Napi::Array result = Napi::Array::New(env);
    int i = 0;
    result.Set(i++, Napi::String::New(env, "RAW"));
    result.Set(i++, Napi::String::New(env, "TEXT"));
    return result;
}

Napi::Value getSupportedJobCommands(const Napi::CallbackInfo& info) 
{
    Napi::Env env = info.Env();
    Napi::Array result = Napi::Array::New(env);
    int i = 0;
    for(StatusMapType::const_iterator itCommand = getJobCommandMap().begin(); itCommand != getJobCommandMap().end(); ++itCommand)
    {
        result.Set(i++, Napi::String::New(env, itCommand->first));
    }
    return result;
}
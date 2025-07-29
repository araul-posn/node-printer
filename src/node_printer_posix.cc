#include "node_printer.hpp"

#include <string>
#include <map>
#include <utility>
#include <sstream>
#include <unistd.h>
#include <stdlib.h>

#include <cups/cups.h>
#include <cups/ppd.h>

#pragma GCC diagnostic ignored "-Wdeprecated-declarations"

namespace
{
    typedef std::map<std::string, int> StatusMapType;
    typedef std::map<std::string, std::string> FormatMapType;

    const StatusMapType& getJobStatusMap()
    {
        static StatusMapType result;
        if(!result.empty())
        {
            return result;
        }
        // add only first time
#define STATUS_PRINTER_ADD(value, type) result.insert(std::make_pair(value, type))
        // Common statuses
        STATUS_PRINTER_ADD("PRINTING", IPP_JOB_PROCESSING);
        STATUS_PRINTER_ADD("PRINTED", IPP_JOB_COMPLETED);
        STATUS_PRINTER_ADD("PAUSED", IPP_JOB_HELD);
        // Specific statuses
        STATUS_PRINTER_ADD("PENDING", IPP_JOB_PENDING);
        STATUS_PRINTER_ADD("PAUSED", IPP_JOB_STOPPED);
        STATUS_PRINTER_ADD("CANCELLED", IPP_JOB_CANCELLED);
        STATUS_PRINTER_ADD("ABORTED", IPP_JOB_ABORTED);

#undef STATUS_PRINTER_ADD
        return result;
    }

    const FormatMapType& getPrinterFormatMap()
    {
        static FormatMapType result;
        if(!result.empty())
        {
            return result;
        }
        result.insert(std::make_pair("RAW", CUPS_FORMAT_RAW));
        result.insert(std::make_pair("TEXT", CUPS_FORMAT_TEXT));
#ifdef CUPS_FORMAT_PDF
        result.insert(std::make_pair("PDF", CUPS_FORMAT_PDF));
#endif
#ifdef CUPS_FORMAT_JPEG
        result.insert(std::make_pair("JPEG", CUPS_FORMAT_JPEG));
#endif
#ifdef CUPS_FORMAT_POSTSCRIPT
        result.insert(std::make_pair("POSTSCRIPT", CUPS_FORMAT_POSTSCRIPT));
#endif
#ifdef CUPS_FORMAT_COMMAND
        result.insert(std::make_pair("COMMAND", CUPS_FORMAT_COMMAND));
#endif
#ifdef CUPS_FORMAT_AUTO
        result.insert(std::make_pair("AUTO", CUPS_FORMAT_AUTO));
#endif
        return result;
    }

    /** Parse job info object.
     * @return error string. if empty, then no error
     */
    std::string parseJobObject(const cups_job_t *job, Napi::Object& result_printer_job, Napi::Env& env)
    {
        //Common fields
        result_printer_job.Set("id", Napi::Number::New(env, job->id));
        result_printer_job.Set("name", Napi::String::New(env, job->title));
        result_printer_job.Set("printerName", Napi::String::New(env, job->dest));
        result_printer_job.Set("user", Napi::String::New(env, job->user));
        std::string job_format(job->format);

        // Try to parse the data format, otherwise will write the unformatted one
        for(FormatMapType::const_iterator itFormat = getPrinterFormatMap().begin(); itFormat != getPrinterFormatMap().end(); ++itFormat)
        {
            if(itFormat->second == job_format)
            {
                job_format = itFormat->first;
                break;
            }
        }

        result_printer_job.Set("format", Napi::String::New(env, job_format.c_str()));
        result_printer_job.Set("priority", Napi::Number::New(env, job->priority));
        result_printer_job.Set("size", Napi::Number::New(env, job->size));
        
        Napi::Array result_printer_job_status = Napi::Array::New(env);
        int i_status = 0;
        for(StatusMapType::const_iterator itStatus = getJobStatusMap().begin(); itStatus != getJobStatusMap().end(); ++itStatus)
        {
            if(job->state == itStatus->second)
            {
                result_printer_job_status.Set(i_status++, Napi::String::New(env, itStatus->first));
            }
        }
        if(i_status == 0)
        {
            // state_reasons is not available in all CUPS versions, use state value instead
            result_printer_job_status.Set(i_status++, Napi::String::New(env, std::to_string(job->state)));
        }
        result_printer_job.Set("status", result_printer_job_status);

        //Specific fields
        result_printer_job.Set("completedTime", Napi::Date::New(env, job->completed_time * 1000));
        result_printer_job.Set("creationTime", Napi::Date::New(env, job->creation_time * 1000));
        result_printer_job.Set("processingTime", Napi::Date::New(env, job->processing_time * 1000));

        // No error
        return "";
    }

    std::string parsePrinterinfo(const cups_dest_t * printer, Napi::Object& result_printer, Napi::Env& env)
    {
        result_printer.Set("name", Napi::String::New(env, printer->name));
        if(printer->instance)
        {
            result_printer.Set("instance", Napi::String::New(env, printer->instance));
        }
        
        result_printer.Set("isDefault", Napi::Boolean::New(env, static_cast<bool>(printer->is_default)));
        
        Napi::Object result_printer_options = Napi::Object::New(env);
        cups_option_t *dest_option = printer->options;
        for(int j = 0; j < printer->num_options; ++j, ++dest_option)
        {
            result_printer_options.Set(dest_option->name, Napi::String::New(env, dest_option->value));
        }
        result_printer.Set("options", result_printer_options);

        return "";
    }
}

Napi::Value getPrinters(const Napi::CallbackInfo& info)
{
    Napi::Env env = info.Env();
    
    cups_dest_t *printers = NULL;
    int printers_size = cupsGetDests(&printers);
    Napi::Array result = Napi::Array::New(env);
    int i = 0;
    cups_dest_t *printer = printers;
    std::string error_str;
    for(; i < printers_size; ++i, ++printer)
    {
        Napi::Object result_printer = Napi::Object::New(env);
        error_str = parsePrinterinfo(printer, result_printer, env);
        if(!error_str.empty())
        {
            goto error;
        }

        // if option is wrong, then the jobs are empty
        if(result_printer.Has("printer-state"))
        {
            // Get printer jobs
            cups_job_t *jobs = NULL;
            int jobs_size = cupsGetJobs(&jobs, printer->name, 0/*0 means all users*/, CUPS_WHICHJOBS_ALL);
            Napi::Array result_priner_jobs = Napi::Array::New(env);
            int j = 0;
            cups_job_t *job = jobs;
            for(; j < jobs_size; ++j, ++job)
            {
                Napi::Object result_printer_job = Napi::Object::New(env);
                error_str = parseJobObject(job, result_printer_job, env);
                if(!error_str.empty())
                {
                    cupsFreeJobs(jobs_size, jobs);
                    goto error;
                }
                result_priner_jobs.Set(j, result_printer_job);
            }
            result_printer.Set("jobs", result_priner_jobs);
            cupsFreeJobs(jobs_size, jobs);
        }
        result.Set(i, result_printer);
    }
    cupsFreeDests(printers_size, printers);
    return result;

error:
    cupsFreeDests(printers_size, printers);
    Napi::Error::New(env, error_str).ThrowAsJavaScriptException();
    return env.Undefined();
}

Napi::Value getDefaultPrinterName(const Napi::CallbackInfo& info) 
{
    Napi::Env env = info.Env();
    const char * printerName = cupsGetDefault();
    return printerName ? Napi::String::New(env, printerName) : env.Undefined();
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
    
    std::string printer_name = info[0].As<Napi::String>().Utf8Value();

    cups_dest_t *printers = NULL, *printer = NULL;
    int printers_size = cupsGetDests(&printers);
    printer = cupsGetDest(printer_name.c_str(), NULL, printers_size, printers);
    Napi::Object result_printer = Napi::Object::New(env);
    if(printer != NULL)
    {
        std::string error_str = parsePrinterinfo(printer, result_printer, env);
        if(!error_str.empty())
        {
            cupsFreeDests(printers_size, printers);
            Napi::Error::New(env, error_str).ThrowAsJavaScriptException();
            return env.Undefined();
        }

        // Get printer jobs
        cups_job_t *jobs = NULL;
        int jobs_size = cupsGetJobs(&jobs, printer->name, 0/*0 means all users*/, CUPS_WHICHJOBS_ALL);
        Napi::Array result_priner_jobs = Napi::Array::New(env);
        int j = 0;
        cups_job_t *job = jobs;
        for(; j < jobs_size; ++j, ++job)
        {
            Napi::Object result_printer_job = Napi::Object::New(env);
            std::string error_str = parseJobObject(job, result_printer_job, env);
            if(!error_str.empty())
            {
                cupsFreeJobs(jobs_size, jobs);
                cupsFreeDests(printers_size, printers);
                Napi::Error::New(env, error_str).ThrowAsJavaScriptException();
                return env.Undefined();
            }
            result_priner_jobs.Set(j, result_printer_job);
        }
        result_printer.Set("jobs", result_priner_jobs);
        cupsFreeJobs(jobs_size, jobs);
    }
    // else printer is not found
    cupsFreeDests(printers_size, printers);
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
    
    std::string printer_name = info[0].As<Napi::String>().Utf8Value();

    // PPD support is deprecated in newer CUPS versions
    // For now, return an empty object
    // TODO: Use cupsCopyDestInfo for modern CUPS API
    Napi::Object ppd_options = Napi::Object::New(env);
    return ppd_options;
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
    
    std::string printer_name = info[0].As<Napi::String>().Utf8Value();
    int job_id = info[1].As<Napi::Number>().Int32Value();
    
    cups_job_t *jobs = NULL;
    int jobs_size = cupsGetJobs(&jobs, printer_name.c_str(), 0/*0 means all users*/, CUPS_WHICHJOBS_ALL);
    Napi::Object result_job;
    int j = 0;
    cups_job_t *job = jobs;
    
    std::string error_str;
    for(; j < jobs_size; ++j, ++job)
    {
        if(job->id != job_id)
        {
            continue;
        }
        result_job = Napi::Object::New(env);
        error_str = parseJobObject(job, result_job, env);
        if(!error_str.empty())
        {
            cupsFreeJobs(jobs_size, jobs);
            Napi::Error::New(env, error_str).ThrowAsJavaScriptException();
            return env.Undefined();
        }
        // Stop search
        break;
    }
    cupsFreeJobs(jobs_size, jobs);
    
    if(!result_job.IsUndefined())
    {
        return result_job;
    }
    
    // return nothing
    return env.Undefined();
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
    
    std::string printer_name = info[0].As<Napi::String>().Utf8Value();
    int job_id = info[1].As<Napi::Number>().Int32Value();
    std::string job_command_str = info[2].As<Napi::String>().Utf8Value();
    
    bool result_ok = false;
    if(job_command_str == "CANCEL")
    {
        result_ok = (cupsCancelJob(printer_name.c_str(), job_id) == 1);
    }
    else
    {
        Napi::Error::New(env, "setJob: unsupported job command. Please see getSupportedJobCommands() for more details").ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    return Napi::Boolean::New(env, result_ok);
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
    
    std::string printer_name;
    // printer name
    if(arg_params.Has("printer"))
    {
        Napi::Value arg_value_printer = arg_params.Get("printer");
        if(!arg_value_printer.IsString())
        {
            Napi::TypeError::New(env, "printDirect:printer parameter must be a string").ThrowAsJavaScriptException();
            return env.Undefined();
        }
        std::string printer_name_str = arg_value_printer.As<Napi::String>().Utf8Value();
        printer_name = printer_name_str;
    }
    else
    {
        // if printer is not specified, then use default printer.
        const char * default_printer_name = cupsGetDefault();
        if(default_printer_name != NULL)
        {
            printer_name = default_printer_name;
        }
    }
    
    // type
    std::string type_str = "RAW";
    if(arg_params.Has("type"))
    {
        Napi::Value arg_value_type = arg_params.Get("type");
        if(!arg_value_type.IsString())
        {
            Napi::TypeError::New(env, "printDirect:type parameter must be a string").ThrowAsJavaScriptException();
            return env.Undefined();
        }
        type_str = arg_value_type.As<Napi::String>().Utf8Value();
    }
    
    // docname
    std::string docname = "node print job";
    if(arg_params.Has("docname"))
    {
        Napi::Value arg_value_docname = arg_params.Get("docname");
        if(!arg_value_docname.IsString())
        {
            Napi::TypeError::New(env, "printDirect:docname parameter must be a string").ThrowAsJavaScriptException();
            return env.Undefined();
        }
        docname = arg_value_docname.As<Napi::String>().Utf8Value();
    }
    
    // options
    cups_option_t *options = NULL;
    int num_options = 0;
    if(arg_params.Has("options"))
    {
        Napi::Value arg_value_options = arg_params.Get("options");
        if(!arg_value_options.IsObject())
        {
            Napi::TypeError::New(env, "printDirect:options parameter must be an object").ThrowAsJavaScriptException();
            return env.Undefined();
        }
        Napi::Object arg_options = arg_value_options.As<Napi::Object>();
        Napi::Array options_names = arg_options.GetPropertyNames();
        
        for(uint32_t i = 0; i < options_names.Length(); ++i)
        {
            Napi::Value name_value = options_names.Get(i);
            if(!name_value.IsString())
            {
                continue;
            }
            std::string name = name_value.As<Napi::String>().Utf8Value();
            Napi::Value option_value = arg_options.Get(name);
            num_options = cupsAddOption(name.c_str(), option_value.ToString().Utf8Value().c_str(), num_options, &options);
        }
    }
    
    FormatMapType::const_iterator itFormat = getPrinterFormatMap().find(type_str);
    if(itFormat == getPrinterFormatMap().end())
    {
        Napi::TypeError::New(env, "printDirect: unsupported format type").ThrowAsJavaScriptException();
        cupsFreeOptions(num_options, options);
        return env.Undefined();
    }
    
    // cupsPrintJob is deprecated, use cupsPrintFile or cupsPrintFiles
    // For direct printing, we need to write data to a temporary file first
    char temp_filename[] = "/tmp/node_printer_XXXXXX";
    int fd = mkstemp(temp_filename);
    if(fd == -1)
    {
        cupsFreeOptions(num_options, options);
        Napi::Error::New(env, "printDirect: failed to create temporary file").ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    if(write(fd, data.c_str(), data.size()) != (ssize_t)data.size())
    {
        close(fd);
        unlink(temp_filename);
        cupsFreeOptions(num_options, options);
        Napi::Error::New(env, "printDirect: failed to write data to temporary file").ThrowAsJavaScriptException();
        return env.Undefined();
    }
    close(fd);
    
    int job_id = cupsPrintFile(printer_name.c_str(), temp_filename, docname.c_str(), num_options, options);
    unlink(temp_filename);
    cupsFreeOptions(num_options, options);
    
    if(job_id == 0)
    {
        std::string error_str = "Print Error: ";
        error_str += cupsLastErrorString();
        Napi::Error::New(env, error_str).ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    Napi::Object result = Napi::Object::New(env);
    result.Set("id", Napi::Number::New(env, job_id));
    
    return result;
}

Napi::Value PrintFile(const Napi::CallbackInfo& info) 
{
    Napi::Env env = info.Env();
    
    if(info.Length() < 1)
    {
        Napi::TypeError::New(env, "printFile:invalid number of arguments (1 expected)").ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    if(!info[0].IsObject())
    {
        Napi::TypeError::New(env, "printFile:first argument must be an object").ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    Napi::Object arg_params = info[0].As<Napi::Object>();
    
    // check filename property
    if(!arg_params.Has("filename"))
    {
        Napi::TypeError::New(env, "printFile:filename parameter is mandatory").ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    // filename
    Napi::Value arg_value_filename = arg_params.Get("filename");
    if(!arg_value_filename.IsString())
    {
        Napi::TypeError::New(env, "printFile:filename parameter must be a string").ThrowAsJavaScriptException();
        return env.Undefined();
    }
    std::string filename = arg_value_filename.As<Napi::String>().Utf8Value();
    
    std::string printer_name;
    // printer name
    if(arg_params.Has("printer"))
    {
        Napi::Value arg_value_printer = arg_params.Get("printer");
        if(!arg_value_printer.IsString())
        {
            Napi::TypeError::New(env, "printFile:printer parameter must be a string").ThrowAsJavaScriptException();
            return env.Undefined();
        }
        printer_name = arg_value_printer.As<Napi::String>().Utf8Value();
    }
    else
    {
        // if printer is not specified, then use default printer.
        const char * default_printer_name = cupsGetDefault();
        if(default_printer_name != NULL)
        {
            printer_name = default_printer_name;
        }
    }
    
    // docname
    std::string title = filename;
    if(arg_params.Has("docname"))
    {
        Napi::Value arg_value_docname = arg_params.Get("docname");
        if(!arg_value_docname.IsString())
        {
            Napi::TypeError::New(env, "printFile:docname parameter must be a string").ThrowAsJavaScriptException();
            return env.Undefined();
        }
        title = arg_value_docname.As<Napi::String>().Utf8Value();
    }
    
    // options
    cups_option_t *options = NULL;
    int num_options = 0;
    if(arg_params.Has("options"))
    {
        Napi::Value arg_value_options = arg_params.Get("options");
        if(!arg_value_options.IsObject())
        {
            Napi::TypeError::New(env, "printFile:options parameter must be an object").ThrowAsJavaScriptException();
            return env.Undefined();
        }
        Napi::Object arg_options = arg_value_options.As<Napi::Object>();
        Napi::Array options_names = arg_options.GetPropertyNames();
        
        for(uint32_t i = 0; i < options_names.Length(); ++i)
        {
            Napi::Value name_value = options_names.Get(i);
            if(!name_value.IsString())
            {
                continue;
            }
            std::string name = name_value.As<Napi::String>().Utf8Value();
            Napi::Value option_value = arg_options.Get(name);
            num_options = cupsAddOption(name.c_str(), option_value.ToString().Utf8Value().c_str(), num_options, &options);
        }
    }
    
    int job_id = cupsPrintFile(printer_name.c_str(), filename.c_str(), title.c_str(), num_options, options);
    cupsFreeOptions(num_options, options);
    
    if(job_id == 0)
    {
        std::string error_str = "Print Error: ";
        error_str += cupsLastErrorString();
        Napi::Error::New(env, error_str).ThrowAsJavaScriptException();
        return env.Undefined();
    }
    
    Napi::Object result = Napi::Object::New(env);
    result.Set("id", Napi::Number::New(env, job_id));
    
    return result;
}

Napi::Value getSupportedPrintFormats(const Napi::CallbackInfo& info)
{
    Napi::Env env = info.Env();
    Napi::Array result = Napi::Array::New(env);
    int i = 0;
    
    for(FormatMapType::const_iterator itFormat = getPrinterFormatMap().begin(); itFormat != getPrinterFormatMap().end(); ++itFormat)
    {
        result.Set(i++, Napi::String::New(env, itFormat->first));
    }
    
    return result;
}

Napi::Value getSupportedJobCommands(const Napi::CallbackInfo& info) 
{
    Napi::Env env = info.Env();
    Napi::Array result = Napi::Array::New(env);
    result.Set((uint32_t)0, Napi::String::New(env, "CANCEL"));
    return result;
}
// BWSSOTestTool.cpp : Defines the entry point for the console application.
//https://support.microsoft.com/en-us/kb/253803
//http://www.codeproject.com/Articles/10236/Database-Development-using-Visual-C-and-OLE-DB-Est

#include "afxwin.h"
#include <atlbase.h>
#include <atldbcli.h>
#include <codecvt>
#include <comdef.h>
#include <fstream>
#include <iomanip>
#include <iostream>
#include <locale>
#include <map>
#include <set>
#include <sstream>
#include <string>
#include <vector>

enum envValueCheckType
{
    NONE,
    EQUAL,
    CONTAINS,
    NONEMPTY
};

enum ConfigSource
{
    SAPLOGON_INI,
    LANDSCAPE_XML,
    UNKNOWN
};

enum AuthType
{
    USERPASS,
    SSO,
    SST,
};

enum warnCondition
{
    SNC_NAME_EMPTY,
    SNC_NOT_ACTIVE,
    SNC_CHOICE_UNKNOWN,
    SLC_NOT_LOGGED_IN,
    LOGONCONFIG_UNKNOWN,
    LOGONCONFIG_ERROR,
    NO_ODBO_PROVIDER,
    NO_32BIT_ODBO_PROVIDER,
    BAD_32BIT_ODBO_PROVIDER_VERSION,
    BAD_SNC_LIB_VERSION_SSO,
    BAD_SNC_LIB_VERSION_SST,
    ENV_VAR_RFC_TRACE_NOT_SET,
    ENV_VAR_APPDATA_NOT_FOUND,
    ENV_VAR_PATH_NOT_FOUND,
    ENV_VAR_PATH_CONTAINS,
    ENV_VAR_SAPLOGON_INI_FILE_NOT_FOUND,
    ENV_VAR_SECUDIR_NOT_FOUND,
    ENV_VAR_SNC_LIB_NOT_FOUND,
    ENV_VAR_SNC_LIB_SSO_EQUALS,
    ENV_VAR_SNC_LIB_SST_EQUALS,
    FILE_NOT_FOUND_SAPLOGON_INI,
    MISSING_SNC_LIB_DLL
};


using namespace std;
IUnknown *	pIUnknown = NULL;
BOOL *		pbValue = FALSE;
CLSID		clsid;
HRESULT		hr;
IDataInitialize*    pIDataInitialize = NULL;
IDBInitialize *		pIDBInitialize = NULL;
IDBProperties*		pIDBProperties = NULL;
wstring_convert<std::codecvt_utf8_utf16<wchar_t>> converter;

wstring sapLogonIniFile = L"";
wstring secudir = L"";
string connectionName = "";
string connectionDefSourceExplicit = "";
wstring sapLogonIniFileExplicit = L"";
set<warnCondition> warningConditions;

// Note: As currently coded, version constants must have all
//       four digits, to prevent uninitialized trailing digits garbage. 
UINT MIN_DRIVER_VERSION_32_BIT[] = { 4,0,0,7 };
UINT MIN_SNC_LIB_VERSION_SST[] = { 8,4,33,0 };
UINT MIN_SNC_LIB_VERSION_SSO[] = { 8,4,49,0 };

const wchar_t* SNC_LIB_SSO = L"C:\\Program Files (x86)\\SAP\\FrontEnd\\SecureLogin\\lib\\sapcrypto.dll";
const wchar_t* SNC_LIB_SST = L"C:\\Program Files\\sap\\crypto\\sapcrypto.dll";

string fileVersionToString( UINT version[4] );

map<ConfigSource, string> configSourceStrings = {
    { SAPLOGON_INI, "saplogon.ini" },
    { LANDSCAPE_XML, "landscape xml" },
    { UNKNOWN, "unknown" }
};

map<warnCondition, string> warnStrings = {
        { SNC_NAME_EMPTY, "SncName is empty" },
        { SNC_NOT_ACTIVE, "SNC is not active" },
        { SNC_CHOICE_UNKNOWN, "SncChoice not recognized" },
        { SLC_NOT_LOGGED_IN, "Not logged in via Secure Logon Client" },
        { LOGONCONFIG_UNKNOWN, "SAP Logon configuration source is not recognized" },
        { LOGONCONFIG_ERROR, "Unable to determine SAP Logon configuration source" },
        { NO_ODBO_PROVIDER, "No ODBO provider found" },
        { NO_32BIT_ODBO_PROVIDER, "No 32-bit ODBO provider found" },
        { BAD_SNC_LIB_VERSION_SSO, "SncLib for SSO version less than " + fileVersionToString(MIN_SNC_LIB_VERSION_SSO) },
        { BAD_SNC_LIB_VERSION_SST, "SncLib for SST version less than " + fileVersionToString(MIN_SNC_LIB_VERSION_SST) },
        { BAD_32BIT_ODBO_PROVIDER_VERSION, "32-bit ODBO provider version less than " + fileVersionToString( MIN_DRIVER_VERSION_32_BIT ) },
        { ENV_VAR_RFC_TRACE_NOT_SET, "Unable to set RFC_TRACE environment variable" },
        { ENV_VAR_APPDATA_NOT_FOUND, "Environment variable APPDATA not found or empty" },
        { ENV_VAR_PATH_NOT_FOUND, "Environment variable Path not found or empty" },
        { ENV_VAR_PATH_CONTAINS, "Environment variable Path does not contain expected substring, e.g., C:\\Program Files (x86)\\SAP\\FrontEnd\\SecureLogin\\lib" },
        { ENV_VAR_SAPLOGON_INI_FILE_NOT_FOUND, "Environment variable SAPLOGON_INI_FILE not found or empty" },
        { ENV_VAR_SECUDIR_NOT_FOUND, "Environment variable SECUDIR not found or empty" },
        { ENV_VAR_SNC_LIB_NOT_FOUND, "Environment variable SNC_LIB not found or empty" },
        { ENV_VAR_SNC_LIB_SSO_EQUALS, "Environment variable SNC_LIB does not have expected value for SSO: " + converter.to_bytes(SNC_LIB_SSO) },
        { ENV_VAR_SNC_LIB_SST_EQUALS, "Environment variable SNC_LIB does not have expected value for SST: " + converter.to_bytes(SNC_LIB_SST) },
        { FILE_NOT_FOUND_SAPLOGON_INI, "saplogon.ini file not found" },
        { MISSING_SNC_LIB_DLL, "SncLib dll not found" }
};

map<AuthType, string> authTypeStrings = {
    { USERPASS, "username/password" },
    { SSO, "SSO" },
    { SST, "impersonate via SST" }
};

string fileVersionToString( UINT version[4] )
{
    string res = "";
    for ( int i = 0; i < 4; ++i )
        res += "." + to_string( version[i] );
    return res.substr( 1 );
}

struct installedComponent {
    string bitness;
    string name;
    string version;
};
vector<installedComponent> vecInstalledComponents;

struct versionedFile
{
    string location;
    UINT version[4];

    versionedFile()
    {
        for ( int i = 0; i < 4; ++i )
            version[i] = 0;
    }

    string versionString() 
    {
        return fileVersionToString( version );
    }

    int compareVersion( UINT other[4] ) const
    {
        for ( int i = 0; i < 4; ++i )
        {
            if ( version[i] < other[i] )
                return -1;
            else if ( version[i] > other[i] )
                return 1;
        }
        return 0; 
    }

    void copyVersion( UINT other[4] )
    {
        for ( int i = 0; i < 4; ++i )
            version[i] = other[i];
    }
};

struct dllStruct
{
    string bitness;
    versionedFile file;
    vector<versionedFile> associatedFiles;
};
vector<dllStruct> vecOdboProviders;
vector<dllStruct> vecSnclibContents;

map<string,string> mapEnvironmentVariables;

ConfigSource logonConfigSource;
string landscapeXMLPath;                // explicit path, if specified, to local xml file
string landscapeXMLLocalFileName;
string landscapeGlobalXMLPath;          // explicit path, if specified, to global xml file
string landscapeGlobalXMLFileName;

struct definedConnectionStruct
{
    string name;
    string sncName;
    string sncChoice;
    string sncNoSSO;
};
definedConnectionStruct selectedConnection;

struct connectionAttemptStruct
{
    string client;
    string language;
    AuthType authType;
    string connectString;
    vector<string> vecMessages;
};
connectionAttemptStruct connectionAttempt;


string serializedJson = "";

/////////////////////////////////// BEGIN PERSISTENCE HACK

string jsonEscape( const string& value )
{
    string retval;

    retval.reserve( value.length() + 32 );

    for ( auto psz = value.cbegin(); psz != value.cend(); ++psz )
    {
        const auto ch = *psz;

        switch ( ch )
        {
        case '\\':
            retval += "\\\\";
            break;

        case '"':
            retval += "\"";
            break;

        case '\n':
            retval += "\\n";
            break;

        case '\t':
            retval += "\\t";
            break;

        default:
            retval += ch;
            break;
        }
    }

    return retval;
}

string jsonName( string name )
{
    return "\"" + name + "\": ";
}
string jsonValue( string value, bool bQuote )
{
    if ( bQuote )
        return "\"" + jsonEscape( value ) + "\"";
    else
        return value;
}
string jsonNameValuePair( string name, string value, bool bQuote = true )
{
    return jsonName( name ) + jsonValue( value, bQuote );
}
string jsonObject( string contents )
{
    return "{" + contents + "}";
}
string jsonArray( string name, string contents )
{
    return jsonName( name ) + "[\n" + contents + "\n]";
}

void stripFinalComma( string& json )
{
    size_t finalCommaPos = json.find_last_of( ",\n" );
    if ( finalCommaPos > 0 )
        json = json.substr( 0, finalCommaPos-1 );
}

void outputInstalledComponents()
{
    cout << "Installed SAP Components" << endl << endl;
    for ( auto ic : vecInstalledComponents )
    {
        cout << "  " << ic.bitness << " bit, " << ic.name << ", version " << ic.version << endl;
    }
    cout << endl << endl;

    string json = "";
    for ( auto ic : vecInstalledComponents )
    {
        string temp = "";
        temp += jsonNameValuePair( "bitness", ic.bitness ) + ", ";
        temp += jsonNameValuePair( "name", ic.name ) + ", ";
        temp += jsonNameValuePair( "version", ic.version );
        temp = jsonObject( temp ) + ",\n";
        json += "  " + temp;
    }
    stripFinalComma( json );
    serializedJson += jsonArray( "installed-components", json ) + ",\n";
}

void outputFileList( const string& caption, const string& jsonArrayName, const vector<dllStruct>& vecFiles)
{
    cout << caption << endl << endl;
    for ( auto op : vecFiles )
    {
        cout << "  " << op.bitness << " bit, " << op.file.location << ", version: " << op.file.versionString() << endl;
        cout << endl;
        for ( auto as : op.associatedFiles )
            cout << "          " << as.location << ", version: " << as.versionString() << endl;
        if ( op.associatedFiles.size() > 0 )
            cout << endl;
    }
    cout << endl << endl;

    string json = "";
    for ( auto op : vecFiles )
    {
        string asJson = "";
        for ( auto as : op.associatedFiles )
        {
            string temp = "";
            temp += jsonNameValuePair( "location", as.location ) + ", ";
            temp += jsonNameValuePair( "version", as.versionString() );
            temp = jsonObject( temp ) + ",\n";
            asJson += "    " + temp;
        }
        stripFinalComma( asJson );

        string temp = "";
        temp += jsonNameValuePair( "bitness", op.bitness ) + ", ";
        temp += jsonNameValuePair( "location", op.file.location ) + ", ";
        temp += jsonNameValuePair( "version", op.file.versionString() ) + ",\n";

        temp += "  " + jsonArray( "associated-files", asJson ) + "\n";

        temp = jsonObject( temp ) + ",\n";
        json += "  " + temp;
    }
    stripFinalComma( json );
    serializedJson += jsonArray( jsonArrayName, json ) + ",\n";
}

void outputOdboProviders()
{
    outputFileList( "SAP ODBO providers", "odbo-providers", vecOdboProviders );
}

void outputSnclibContents()
{
    outputFileList("Snc Lib", "snc-lib", vecSnclibContents );
}

void outputEnvironmentVariables()
{
    cout << "Environment Variables" << endl << endl;
    for ( auto it = mapEnvironmentVariables.begin(); it != mapEnvironmentVariables.end(); ++it )
    {
        cout << "  " << it->first << "=" << it->second << endl;
    }
    cout << endl << endl;

    string json = "";
    for ( auto it = mapEnvironmentVariables.begin(); it != mapEnvironmentVariables.end(); ++it )
    {
        string temp = "";
        temp += jsonNameValuePair( "name", it->first ) + ", ";
        temp += jsonNameValuePair( "value", it->second );
        temp = jsonObject( temp ) + ",\n";
        json += "  " + temp;
    }
    stripFinalComma( json );
    serializedJson += jsonArray( "environment-variables", json ) + ",\n";
}

void configPathsAndSources(ConfigSource configSource, string& path, string& pathGlobal, string& pathSource, string& pathGlobalSource)
{
    switch ( configSource )
    {
    case ConfigSource::LANDSCAPE_XML:
        path = landscapeXMLLocalFileName;
        pathSource += landscapeXMLPath.empty() ? "default" : "PathConfigFilesLocal";
        pathGlobal = landscapeGlobalXMLFileName;
        pathGlobalSource += landscapeGlobalXMLPath.empty() ? "default" : "LandscapeFileOnServer";
        break;
    case ConfigSource::SAPLOGON_INI:
        path = converter.to_bytes(sapLogonIniFile);
        pathSource = sapLogonIniFileExplicit.empty() ? "default" : "explicit";
        break;
    default:
        break;
    }
}

void outputConfigPathsAndSources(ConfigSource configSource, const string& path, const string& pathGlobal, const string& pathSource, const string& pathGlobalSource)
{
    switch ( configSource )
    {
    case ConfigSource::LANDSCAPE_XML:
        cout << "  landscapeXMLPath:       " << path << " (" << pathSource << ")" << endl;
        cout << "  landscapeGlobalXMLPath: " << pathGlobal << " (" << pathGlobalSource << ")" << endl;
        break;
    case ConfigSource::SAPLOGON_INI:
        cout << "  sapLogonIniFile:        " << path << " (" << pathSource << ")" << endl;
        break;
    default:
        break;
    }
    cout << endl << endl;
}

void outputLogonConfigSource()
{
    string path;
    string pathGlobal;
    string pathSource;
    string pathGlobalSource;

    cout << "SAP Logon configuration source: " << configSourceStrings[logonConfigSource] << endl;
    configPathsAndSources(logonConfigSource, path, pathGlobal, pathSource, pathGlobalSource);
    outputConfigPathsAndSources(logonConfigSource, path, pathGlobal, pathSource, pathGlobalSource);

    string temp = "";
    temp += jsonNameValuePair( "source", configSourceStrings[logonConfigSource] ) + ", ";
    temp += jsonNameValuePair( "path", path ) + ", ";
    temp += jsonNameValuePair( "path-source", pathSource);
    if ( !pathGlobal.empty() )
    {
        temp += ", ";
        temp += jsonNameValuePair( "path-global", pathGlobal ) + ", ";
        temp += jsonNameValuePair( "path-global-source", pathGlobalSource );
    }
    temp = jsonObject( temp );

    serializedJson += jsonNameValuePair( "logon-config-source", temp, false ) + ",\n";
}

void outputSelectedConnection()
{
    string sncChoice = selectedConnection.sncChoice;
    string sncChoiceDesc;
    if ( sncChoice == "" || sncChoice == "0" || sncChoice == "-1" ) {
        // Don't currently understand the difference between -1 and 0. I have two connections whose property
        // tabs look identical in SAP Logon, but one has sapChoice 0 and the other -1. 
        sncChoiceDesc = "Not Active";
    }
    else if ( sncChoice == "1" )
        sncChoiceDesc = "Authentication only";
    else if ( sncChoice == "2" )
        sncChoiceDesc = "Integrity protection";
    else if ( sncChoice == "3" )
        sncChoiceDesc = "Privacy protection";
    else if ( sncChoice == "9" )
        sncChoiceDesc = "Maximum security settings available";
    else
        sncChoiceDesc = "Not recognized";

    cout << "Selected Connection" << endl;
    cout << "  Name:      " << selectedConnection.name << endl;
    cout << "  SncName:   " << selectedConnection.sncName << endl;
    cout << "  SncChoice: " << selectedConnection.sncChoice + "  " + sncChoiceDesc << endl;
    cout << "  SncNoSSO:  " << selectedConnection.sncNoSSO << endl;
    cout << endl << endl;

    string temp = "";
    temp += jsonNameValuePair( "name", selectedConnection.name ) + ", ";
    temp += jsonNameValuePair( "SncName", selectedConnection.sncName ) + ", ";
    temp += jsonNameValuePair( "SncChoice", selectedConnection.sncChoice + "  " + sncChoiceDesc ) + ", ";
    temp += jsonNameValuePair( "SncNoSSO", selectedConnection.sncNoSSO );
    temp = jsonObject( temp );

    serializedJson += jsonNameValuePair( "selected-connection", temp, false) + ",\n";
}

void outputConnectionAttempt()
{
    cout << endl << endl << "Connecting..." << endl;
    for ( auto m : connectionAttempt.vecMessages )
        cout << "  " << m << endl;
    cout << endl << endl;

    string tempMessages = "";
    {
        string temp = "";
        for ( auto m : connectionAttempt.vecMessages )
            temp += "    " + jsonValue( m, true ) + ",\n";
        stripFinalComma( temp );
        tempMessages = "  " + jsonArray( "messages", temp ) + "\n";
    }

    string temp = "\n";
    temp += "  " + jsonNameValuePair( "client", connectionAttempt.client ) + ",\n";
    temp += "  " + jsonNameValuePair( "language", connectionAttempt.language ) + ",\n";
    temp += "  " + jsonNameValuePair( "auth-type", authTypeStrings[connectionAttempt.authType] ) + ",\n";
    temp += "  " + jsonNameValuePair( "connectString", connectionAttempt.connectString ) + ",\n";
    temp += tempMessages;
    temp = jsonObject( temp );
    serializedJson += jsonNameValuePair( "connection-attempt", temp, false ) + ",\n";
}

// 
// For SNC conditions only display the warning if bSSOConnect is true
//
void outputConfigWarnings( bool bSSOConnect, bool bImpersonateViaSST )
{
    vector<string> warnings;

    for ( auto warnCondition : warningConditions )
    {
        switch ( warnCondition )
        {
        case ENV_VAR_PATH_NOT_FOUND:
        case ENV_VAR_PATH_CONTAINS:
        case SLC_NOT_LOGGED_IN:
        case BAD_SNC_LIB_VERSION_SSO:
            if ( bSSOConnect )
                warnings.push_back( warnStrings[warnCondition] );
            break;
        case BAD_SNC_LIB_VERSION_SST:
            if ( bImpersonateViaSST )
                warnings.push_back(warnStrings[warnCondition]);
            break;
        case ENV_VAR_SNC_LIB_NOT_FOUND:
        case ENV_VAR_SNC_LIB_SSO_EQUALS:
        case ENV_VAR_SNC_LIB_SST_EQUALS:
        case SNC_NAME_EMPTY:
        case SNC_NOT_ACTIVE:
        case SNC_CHOICE_UNKNOWN:
        case NO_32BIT_ODBO_PROVIDER:
        case BAD_32BIT_ODBO_PROVIDER_VERSION:
            if ( bSSOConnect || bImpersonateViaSST )
                warnings.push_back(warnStrings[warnCondition]);
            break;
        case ENV_VAR_APPDATA_NOT_FOUND:
            if ( logonConfigSource == ConfigSource::LANDSCAPE_XML )
                warnings.push_back( warnStrings[warnCondition] );
            break;
        default:
            warnings.push_back( warnStrings[warnCondition] );
            break;
        }
    }

    cout << endl;
    if ( warnings.size() == 0 )
        cout << "No Configuration Warnings" << endl;
    else
    {
        cout << "Configuration Warnings:" << endl;
        for ( auto warn : warnings )
            cout << "  " << warn << endl;

        cout << endl;
    }

    string tempMessages = "";
    {
        string temp = "";
        for ( auto warn : warnings )
            temp += "  " + jsonValue( warn, true ) + ",\n";
        stripFinalComma( temp );
        tempMessages = jsonArray( "configuration-warnings", temp );
    }
    serializedJson += tempMessages + ",\n";
}

void outputExitCode(int exitCode)
{
    serializedJson += jsonNameValuePair( "exit-code", to_string( exitCode ) );
}

void persistResults()
{
    serializedJson = "{\n" + serializedJson + "\n}";

    ofstream f;
    f.open( "BWSSOresults.json" );
    f << serializedJson;
    f.close();
}

/////////////////////////////////// END PERSISTENCE HACK


void printHResult(HRESULT hr) {
    _com_error err(hr);
    auto errMsg8 = converter.to_bytes(err.ErrorMessage());
    connectionAttempt.vecMessages.push_back( "Error: " + errMsg8 );
}

int checkHResult(char * desc, HRESULT hr) {
    if (SUCCEEDED(hr)) {
        connectionAttempt.vecMessages.push_back( string(desc) + " succeeded" );
        return 0;
    }
    else {
        connectionAttempt.vecMessages.push_back( string( desc ) + " failed" );
        printHResult( hr );
        return 1;
    }
}

void initialize( bool bImpersonate ) {

    HRESULT hr;
    
    if ( bImpersonate )
    {
        hr = CoInitializeEx( 0, COINIT_MULTITHREADED );
        if ( checkHResult( "CoInitializeEx", hr ) ) return;

        hr = CoInitializeSecurity(
            NULL,
            -1,                          // COM negotiates service
            NULL,                        // Authentication services
            NULL,                        // Reserved
            RPC_C_AUTHN_LEVEL_DEFAULT,   // Default authentication 
            RPC_C_IMP_LEVEL_IMPERSONATE, // Default Impersonation  
            NULL,                        // Authentication info
            EOAC_NONE,                   // Additional capabilities 
            NULL                         // Reserved
            );
        if ( checkHResult( "CoInitializeSecurity", hr ) ) return;
    }
    else
    {
        hr = CoInitialize(NULL);
        if ( checkHResult( "CoInitialize", hr ) ) return;
    }

    hr = CoCreateInstance( CLSID_MSDAINITIALIZE, NULL, CLSCTX_INPROC_SERVER, IID_IDataInitialize, (void**)&pIDataInitialize );
}

vector<wstring> getProviders(boolean list) {
    CEnumerator oProviders;
    HRESULT hr = oProviders.Open();
    vector<wstring> providers = vector < wstring >() ;

    if (SUCCEEDED(hr))
    {
        // The following macro is to initialize
        // the conversion routines
        USES_CONVERSION;
        if (list) cout << "Provider\tDescription" << endl;
        while (oProviders.MoveNext() == S_OK)
        {
            if(list) cout << W2A(oProviders.m_szName) << "\t" << W2A(oProviders.m_szDescription) << endl;
            providers.push_back( converter.from_bytes( W2A( oProviders.m_szName ) ) );
        }
        oProviders.Close();
    } else {
        printHResult(hr);
    }
    return providers;
}

string trimNulls( wstring s )
{
    char nullChar = 0;
    size_t nullPos = s.find_first_of( nullChar );
    if ( nullPos >= 0 )
        s = s.substr( 0, nullPos );

    return converter.to_bytes(s);
}

void setEnvVariable( LPCWSTR key, LPCWSTR value, warnCondition wc )
{
    if ( !SetEnvironmentVariableW( key, value ) )
    {
        warningConditions.insert( wc );
    }
}

void getEnvVariable( LPCWSTR key, bool& success, wstring& value )
{
    WCHAR actualVal[32767] = L"";
    DWORD res = 0;

    res = GetEnvironmentVariableW( key, actualVal, 32767 );
    success = res != 0;
    value = actualVal;

    mapEnvironmentVariables[trimNulls( key )] = trimNulls( actualVal );
}

void environmentVariables( bool bImpersonateViaSST )
{
    setEnvVariable( L"RFC_TRACE", L"1", warnCondition::ENV_VAR_RFC_TRACE_NOT_SET );

    bool success = false;
    wstring value = L"";

    if ( sapLogonIniFileExplicit != L"" )
        sapLogonIniFile = sapLogonIniFileExplicit;
    else
    {
        getEnvVariable( L"SAPLOGON_INI_FILE", success, sapLogonIniFile );
        if ( !success || sapLogonIniFile == L"" )
            warningConditions.insert( ENV_VAR_SAPLOGON_INI_FILE_NOT_FOUND );
    }

    getEnvVariable( L"SNC_LIB", success, value );
    if ( !success || value == L"" )
        warningConditions.insert( ENV_VAR_SNC_LIB_NOT_FOUND );
    else
    {
        if ( bImpersonateViaSST )
        {
            if ( _wcsicmp(value.c_str(), SNC_LIB_SST) != 0 )
                warningConditions.insert(ENV_VAR_SNC_LIB_SST_EQUALS);
        }
        else
        {
            if ( _wcsicmp(value.c_str(), SNC_LIB_SSO) != 0 )
                warningConditions.insert(ENV_VAR_SNC_LIB_SSO_EQUALS);
        }
    }

    if ( bImpersonateViaSST )
    {
        getEnvVariable( L"SECUDIR", success, secudir );
        if ( !success || value == L"" )
            warningConditions.insert(ENV_VAR_SECUDIR_NOT_FOUND);
    }

    getEnvVariable( L"APPDATA", success, value );
    if ( !success || value == L"" )
        warningConditions.insert( ENV_VAR_APPDATA_NOT_FOUND );

    outputEnvironmentVariables();
}

void DisplayGetPrivateProfileStringError( wstring step )
{
    DWORD errCode = GetLastError();
    if ( errCode != ERROR_FILE_NOT_FOUND )
    {
        wcout << "profile error (not found) at " << step << ":" << errCode << endl;
    }
    if ( errCode != ERROR_SUCCESS )
        wcout << "profile error (non success) at " << step << ":" << errCode << endl;
}

definedConnectionStruct getIniProperties(const wstring& wstrItem)
{
    // TODO reconsider warnings, since now we may read config data for *every* connection

    definedConnectionStruct connection;

    WCHAR sysName[32767] = L"";
    DWORD nameLength = _countof( sysName );

    memset( sysName, 0, sizeof( sysName ) );

    if ( GetPrivateProfileString( L"Description", wstrItem.c_str(), L"", sysName, nameLength, sapLogonIniFile.c_str() ) == 0 ) {
        DWORD errCode = GetLastError();
        if ( errCode != ERROR_FILE_NOT_FOUND )
        {
            cout << "error enumerating defined connections";
            //LogMessage( TFormatString( TS( "GetPrivateProfileString(%1): error=%2" ) ).arg( item ).arg( errCode ) );
        }
        if ( errCode != ERROR_SUCCESS )
            return connection;
    }

    connection.name = trimNulls( sysName );

    wstring step;

    // SncName

    memset( sysName, 0, sizeof( sysName ) );

    if ( GetPrivateProfileString( L"SncName", wstrItem.c_str(), L"", sysName, nameLength, sapLogonIniFile.c_str() ) == 0 ) {
        //Reach here if there is no SncName entry; looks like SAP Logon won't let you delete an SncName once added,
        //but for testing purposes you can just edit the saplogon.ini file to empty out the value. 
        //DisplayGetPrivateProfileStringError(step);

        //warningConditions.insert( SNC_NAME_EMPTY );
    }
    else
    {
        wstring sncName = sysName;
        connection.sncName = trimNulls( sncName );
        //if ( sncName == L"" )
        //{
        //    warningConditions.insert( SNC_NAME_EMPTY );
        //}
    }

    // SncChoice

    memset( sysName, 0, sizeof( sysName ) );

    if ( GetPrivateProfileString( L"SncChoice", wstrItem.c_str(), L"", sysName, nameLength, sapLogonIniFile.c_str() ) == 0 ) {
        DisplayGetPrivateProfileStringError( step );
    }
    else
    {
        wstring sncChoice = sysName;
        connection.sncChoice = trimNulls( sncChoice );
        //if ( sncChoice == L"" || sncChoice == L"0" || sncChoice == L"-1" ) {
        //    // Don't currently understand the difference between -1 and 0. I have two connections whose property
        //    // tabs look identical in SAP Logon, but one has sapChoice 0 and the other -1. 
        //    warningConditions.insert( SNC_NOT_ACTIVE );
        //}
        //else if ( (sncChoice != L"1") && (sncChoice != L"2") && (sncChoice != L"3") && (sncChoice != L"9") ) {
        //    warningConditions.insert( SNC_CHOICE_UNKNOWN );
        //}
    }

    // SncNoSSO

    memset( sysName, 0, sizeof( sysName ) );

    if ( GetPrivateProfileString( L"SncNoSSO", wstrItem.c_str(), L"", sysName, nameLength, sapLogonIniFile.c_str() ) == 0 ) {
        DisplayGetPrivateProfileStringError( step );
    }
    else
    {
        wstring sncNoSSO = sysName;
        connection.sncNoSSO = trimNulls( sncNoSSO );
    }

    return connection;
}

map<string,definedConnectionStruct> connectionDefsFromIni()
{
    map<string, definedConnectionStruct> definedConnections;

    if ( sapLogonIniFile == L"" )
        return definedConnections;

    ifstream infile(sapLogonIniFile);
    if ( !infile.good() )
    {
        warningConditions.insert(FILE_NOT_FOUND_SAPLOGON_INI);
        return definedConnections;
    }

    for ( size_t index = 1; ; ++index ) {

        std::wostringstream ws;
        ws << L"Item";
        ws << index;
        wstring wstrItem = ws.str();

        auto connection = getIniProperties(wstrItem);
        if ( connection.name == "" )  // TODO works for ini; not sure best approach
            break;

        // TODO don't present empty strings as connections.
        definedConnections[ connection.name ] = connection;
    }

    return definedConnections;
}

/////////////////////////////////// BEGIN XML HACK

string GetServicesBody( const string& xml )
{
    string tagStart = "<Services>";
    string tagEnd = "</Services>";

    size_t idxStart;
    if ( (idxStart = xml.find( tagStart )) == std::string::npos )
        return "";
    idxStart += tagStart.length();

    size_t idxEnd;
    if ( (idxEnd = xml.find( tagEnd, idxStart )) == std::string::npos )
        return "";

    return xml.substr( idxStart, idxEnd - idxStart );
}

definedConnectionStruct getServiceXMLAttributes( const string& serviceXML )
{
    definedConnectionStruct connection;

    string attrName = " type=\"";

    size_t idxStart = serviceXML.find( attrName );
    if ( idxStart == std::string::npos )
        return connection;

    idxStart += attrName.length();

    size_t idxEnd = serviceXML.find( "\"", idxStart );
    if ( idxEnd == std::string::npos )
        return connection;

    string strType = serviceXML.substr( idxStart, idxEnd - idxStart );
    if ( strType != "SAPGUI" )
        return connection;

    // ***
    attrName = " name=\"";

    idxStart = serviceXML.find( attrName );
    if ( idxStart == std::string::npos )
        return connection;

    idxStart += attrName.length();

    idxEnd = serviceXML.find( "\"", idxStart );
    if ( idxEnd == std::string::npos )
        return connection;

    connection.name = serviceXML.substr( idxStart, idxEnd - idxStart );

    // ***
    attrName = " sncname=\"";

    idxStart = serviceXML.find( attrName );
    if ( idxStart != std::string::npos )
    {
        idxStart += attrName.length();

        idxEnd = serviceXML.find( "\"", idxStart );
        if ( idxEnd != std::string::npos )
            connection.sncName = serviceXML.substr( idxStart, idxEnd - idxStart );
    }

    // ***
    attrName = " sncop=\"";

    connection.sncChoice = "0"; // default, TODO confirm

    idxStart = serviceXML.find( attrName );
    if ( idxStart != std::string::npos )
    {
        idxStart += attrName.length();

        idxEnd = serviceXML.find( "\"", idxStart );
        if ( idxEnd != std::string::npos )
            connection.sncChoice = serviceXML.substr( idxStart, idxEnd - idxStart );
    }

    return connection;
}

void connectionDefsFromXMLString( map<string, definedConnectionStruct>& connectionDefs, const string& xml )
{
    string servicesXML = GetServicesBody( xml );

    string tagStart = "<Service ";
    string tagEnd = "/>";

    size_t idxStart = 0;
    while ( (idxStart = servicesXML.find( tagStart, idxStart )) != std::string::npos )
    {
        idxStart += tagStart.length() - 1;  // -1 because we should leave the whitespace with this bogus approach

        size_t idxEnd;
        if ( (idxEnd = servicesXML.find( tagEnd, idxStart )) == std::string::npos )
            break;

        string serviceXML = servicesXML.substr( idxStart, idxEnd - idxStart );

        auto connection = getServiceXMLAttributes( serviceXML );

        idxStart = idxEnd; // adjust

        if ( connection.name == "" )  // TODO works for ini; not sure best approach
            continue;

        // TODO don't present empty strings as connections.
        connectionDefs[connection.name] = connection;
    }
}

/////////////////////////////////// END XML HACK

string XMLFileContents( const string& fileName )
{
    stringstream buff;

    ifstream infile( fileName );
    if ( !infile.is_open() )
        return "";

    buff << infile.rdbuf();
    infile.close();

    string myStr = buff.str();
    return myStr;
}

map<string, definedConnectionStruct>  connectionDefsFromXML()
{
    map<string, definedConnectionStruct> definedConnections;

    string appData;
    {
        auto it = mapEnvironmentVariables.find( "APPDATA" );
        if ( it != mapEnvironmentVariables.end() )
            appData = it->second;
        else
            ; // TODO something reasonable
    }

    const string xmlPathDefault = appData + "\\SAP\\Common";

    landscapeXMLLocalFileName = ( !landscapeXMLPath.empty() ? landscapeXMLPath : xmlPathDefault ) + "\\SAPUILandscape.xml";

    // NOTE: we do landscape xml first, then global landscape, because
    // connectionDefsFromXML string overwrites same-named connection defs, 
    // and we want global entries to "win" if there are dupes

    string xmlStr = XMLFileContents(landscapeXMLLocalFileName);
    if ( !xmlStr.empty() )
        connectionDefsFromXMLString( definedConnections, xmlStr );
    else
        ; // TODO warn

    landscapeGlobalXMLFileName = !landscapeGlobalXMLPath.empty() ? landscapeGlobalXMLPath : xmlPathDefault + "\\SAPUILandscapeGlobal.xml";

    xmlStr = XMLFileContents(landscapeGlobalXMLFileName);
    if ( !xmlStr.empty() )
        connectionDefsFromXMLString( definedConnections, xmlStr );
    else
        ; // TODO warn

    return definedConnections;
}


void checkConnectionWarnings( const definedConnectionStruct& connection )
{
    if ( connection.sncName == "" )
        warningConditions.insert( SNC_NAME_EMPTY );

    if ( connection.sncChoice == "" || connection.sncChoice == "0" || connection.sncChoice == "-1" ) {
        // Don't currently understand the difference between -1 and 0. I have two connections whose property
        // tabs look identical in SAP Logon, but one has sapChoice 0 and the other -1. 
        warningConditions.insert( SNC_NOT_ACTIVE );
    }
    else if ( (connection.sncChoice != "1") && (connection.sncChoice != "2") && (connection.sncChoice != "3") && (connection.sncChoice != "9") ) {
        warningConditions.insert( SNC_CHOICE_UNKNOWN );
    }
}

// TODO 
//  1) fix error message
//  2) currently does not suppress connections with empty names. 
void chooseConnection()
{
    map<string, definedConnectionStruct> definedConnections;

    if ( logonConfigSource == ConfigSource::SAPLOGON_INI )
        definedConnections = connectionDefsFromIni();
    else if ( logonConfigSource == ConfigSource::LANDSCAPE_XML )
        definedConnections = connectionDefsFromXML();
    else
        return;

    // TODO may want to move this
    outputLogonConfigSource();

    if ( connectionName == "" )
    {
        cout << "Defined Connections:\n";

        if ( definedConnections.size() == 0 )
        {
            cout << endl << endl;
            return;
        }

        size_t index = 0;
        vector<string> connectionNameVec;
        for ( auto it = definedConnections.begin(); it != definedConnections.end(); ++it )
        {
            string connName = it->first;
            connectionNameVec.push_back(connName);

            index++;
            cout << "  " << index << "  " << connName << "\n";
        }

        string strIdx;
        string promptRange = (definedConnections.size() == 1) ? " (1)" : ", 1-" + to_string( definedConnections.size() );
        cout << "  Enter connection to use" << promptRange << ": ";
        getline( cin, strIdx );
        cout << endl;

        if ( definedConnections.size() == 1 && strIdx == "" )
            strIdx = "1";

        size_t selectedIdx;
        istringstream iss( strIdx );
        if ( iss >> selectedIdx && iss.eof() && 1 <= selectedIdx && selectedIdx <= definedConnections.size() )
        {
            connectionName = connectionNameVec[selectedIdx - 1];
        }
        else
        {
            cout << "Invalid entry" << endl << endl;
            return;
        }
    }

    auto itSelectedConnection = definedConnections.find( connectionName );

    // Exit if entered connection name is invalid
    if ( itSelectedConnection == definedConnections.end() )
    {
        cout << "Unable to find connection '" << connectionName << "', quitting." << endl << endl;
        exit(1);
    }
    selectedConnection = itSelectedConnection->second;

    checkConnectionWarnings( selectedConnection );

    outputSelectedConnection();
}

wstring RegistryQueryValue( HKEY hKey,
    LPCTSTR szName )
{
    wstring value;
    DWORD dwType;
    DWORD dwSize = 0;

    if ( ::RegQueryValueEx( hKey, szName, NULL, &dwType, NULL, &dwSize ) == ERROR_SUCCESS && dwSize > 0 )
    {
        value.resize( dwSize );

        ::RegQueryValueEx( hKey, szName, NULL, &dwType, (LPBYTE)&value[0], &dwSize );
    }

    return value;
}

void RegistryEnum(bool b64bit)
{
    wstring bitness;

    // simplistic way to access both 32 bit and 64 bit registry entries

    REGSAM sam;
    if ( b64bit )
    {
        bitness = L"64";
        sam = KEY_READ | KEY_WOW64_64KEY;
    }
    else
    {
        bitness = L"32";
        sam = KEY_READ | KEY_WOW64_32KEY;
    }

    HKEY hKey;
    LONG ret = ::RegOpenKeyEx(
        HKEY_LOCAL_MACHINE,
        _T( "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall" ),
        0,
        sam,
        &hKey
        );

    if ( ret != ERROR_SUCCESS )
        return;

    DWORD dwIndex = 0;
    DWORD cbName = 1024;
    TCHAR szSubKeyName[1024];

    installedComponent ic;
    ic.bitness = converter.to_bytes(bitness);

    while ( (ret = ::RegEnumKeyEx(
        hKey,
        dwIndex,
        szSubKeyName,
        &cbName,
        NULL,
        NULL,
        NULL,
        NULL )) != ERROR_NO_MORE_ITEMS )
    {
        if ( ret == ERROR_SUCCESS )
        {
            HKEY hItem;
            if ( ::RegOpenKeyEx( hKey, szSubKeyName, 0, sam, &hItem ) != ERROR_SUCCESS ) //KEY_READ
                continue;

            wstring name = RegistryQueryValue( hItem, _T( "DisplayName" ) );

            if ( name.substr( 0, 3 ) == L"SAP" )
            {
                wstring version = RegistryQueryValue( hItem, _T( "DisplayVersion" ) );
                //wstring publisher = RegistryQueryValue( hItem, _T( "Publisher" ) );
                //wstring location = RegistryQueryValue( hItem, _T( "InstallLocation" ) );

                char nullChar = 0;
                size_t nameLength = name.find_first_of( nullChar );
                size_t versionLength = version.find_first_of( nullChar );
                //ic.name = converter.to_bytes(name.substr( 0, nameLength ));
                //ic.version = converter.to_bytes( version.substr(0, versionLength ));
                ic.name = trimNulls( name );
                ic.version = trimNulls( version );
                vecInstalledComponents.push_back( ic );
            }

            ::RegCloseKey( hItem );
        }
        dwIndex++;
        cbName = 1024;
    }
    ::RegCloseKey( hKey );
}

void installedComponents()
{
    RegistryEnum( true );
    RegistryEnum( false );

    outputInstalledComponents();
}

void locateLandscapeXMLFiles()
{
    HKEY hkey;
    LPCWSTR key = L"Software\\SAP\\SAPLogon\\Options";

    if ( RegOpenKeyExW( HKEY_CURRENT_USER, key, 0, KEY_READ, &hkey ) == ERROR_SUCCESS )
    {
        // landscapeXMLPath - explicit path to local xml file
        //   1. PathConfigFilesLocal
        // If not found, connectionDefsFromXML() defaults xml file location to APPDATA + "\\SAP\\Common" 
        {
            TCHAR buf[512];
            const int bufSize = sizeof( buf );

            if ( RegQueryValueExW( hkey, L"PathConfigFilesLocal", 0, nullptr, (LPBYTE)buf, (LPDWORD)&bufSize ) ==
                ERROR_SUCCESS )
            {
                landscapeXMLPath = converter.to_bytes( buf );
            }
            // TODO want any output if not success? 
        }

        // landscapeGlobalXMLPath - explicit path to global xml file
        //   1. LandscapeFileOnServer registry entry
        //   2. CoreLandscapeFileOnServer registry entry
        //   3. PathConfigFilesLocal + \\SAPUILandscapeGlobal.xml
        // If not found, connectionDefsFromXML() defaults xml file location to APPDATA + "\\SAP\\Common\\SAPUILandscapeGlobal.xml" 
        {
            TCHAR buf[512];
            const int bufSize = sizeof(buf);

            if ( RegQueryValueExW(hkey, L"LandscapeFileOnServer", 0, nullptr, (LPBYTE)buf, (LPDWORD)&bufSize) ==
                ERROR_SUCCESS )
            {
                landscapeGlobalXMLPath = converter.to_bytes(buf);
            }
        }
        if ( landscapeGlobalXMLPath.empty() )
        {
            TCHAR buf[512];
            const int bufSize = sizeof(buf);

            if ( RegQueryValueExW(hkey, L"CoreLandscapeFileOnServer", 0, nullptr, (LPBYTE)buf, (LPDWORD)&bufSize) ==
                ERROR_SUCCESS )
            {
                landscapeGlobalXMLPath = converter.to_bytes(buf);
            }
        }
        if ( landscapeGlobalXMLPath.empty() && !landscapeXMLPath.empty() )
        {
            landscapeGlobalXMLPath = landscapeXMLPath + "\\SAPUILandscapeGlobal.xml";
        }

        RegCloseKey( hkey );
    }
}

void sapLogonConfigSource()
{
    if ( connectionDefSourceExplicit == "i" )
        logonConfigSource = ConfigSource::SAPLOGON_INI;
    else if ( connectionDefSourceExplicit == "l" )
        logonConfigSource = ConfigSource::LANDSCAPE_XML;
    else
    {
        const wstring warnStr = L"Unable to determine SAP Logon configuration source";

        HKEY hkey;
        LPCWSTR key = L"SOFTWARE\\Wow6432Node\\SAP\\SAPLogon";
        if ( RegOpenKeyExW( HKEY_LOCAL_MACHINE, key, 0, KEY_READ, &hkey ) == ERROR_SUCCESS )
        {
            TCHAR buf[512];
            const int bufSize = sizeof( buf );

            if ( RegQueryValueExW( hkey, L"LandscapeFormatEnabled", 0, nullptr, (LPBYTE)buf, (LPDWORD)&bufSize ) ==
                ERROR_SUCCESS )
            {
                switch ( buf[0] )
                {
                case 0x0:
                    logonConfigSource = ConfigSource::SAPLOGON_INI;
                    break;
                case 0x1:
                    //warningConditions.insert( LOGONCONFIG_XML );
                    logonConfigSource = ConfigSource::LANDSCAPE_XML;
                    break;
                default:
                    warningConditions.insert( LOGONCONFIG_UNKNOWN );
                    logonConfigSource = ConfigSource::UNKNOWN;
                }
                wcout << endl;
            }
            else
            {
                // saplogin.ini is used if the registry entry does not exist
                logonConfigSource = ConfigSource::SAPLOGON_INI;
            }

            RegCloseKey( hkey );
        }
        else
        {
            // saplogin.ini is used if the registry entry does not exist
            logonConfigSource = ConfigSource::SAPLOGON_INI;
        }
    }

    if ( logonConfigSource == ConfigSource::LANDSCAPE_XML )
    {
        locateLandscapeXMLFiles();
    }

    // moved, for now, to chooseConnection()
    //outputLogonConfigSource();  
}

string getPassword()
{
    string res;

    // set console mode to no echo
    DWORD mode = 0;
    HANDLE hin = GetStdHandle( STD_INPUT_HANDLE );
    SetConsoleMode( hin, mode & ~(ENABLE_ECHO_INPUT | ENABLE_LINE_INPUT) );

    string prompt = "Enter password: ";
    DWORD count = 0;
    HANDLE hout = GetStdHandle( STD_OUTPUT_HANDLE );
    WriteConsoleA( hout, prompt.c_str(), prompt.length(), &count, NULL );
    char c;
    while ( ReadConsoleA( hin, &c, 1, &count, NULL ) && (c != '\r') && (c != '\n') )
    {
        if ( c == '\b' )
        {
            if ( res.length() )
            {
                WriteConsoleA( hout, "\b \b", 3, &count, NULL );
                res.erase( res.end() - 1 );
            }
        }
        else
        {
            WriteConsoleA( hout, "*", 1, &count, NULL );
            res.push_back( c );
        }
    }

    // restore console mode
    SetConsoleMode( hin, mode );

    return res;
}


wstring getConnectString( bool& bSSOConnect, bool bImpersonateViaSST )
{
    const wstring provider = L"MDrmSap.2";

    wstring tempStr = L"";
    wstring client = L"800";
    wstring language = L"";
    wstring username = L"developer3";
    wstring password = L"";

    wcout << "Enter client (" << client << "): ";
    getline( wcin, tempStr );
    if ( tempStr != L"" )
        client = tempStr;

    wcout << "Enter language (" << language << "): ";
    getline( wcin, tempStr );
    if ( tempStr != L"" )
        language = tempStr;

    if ( !bImpersonateViaSST )
    {
        wcout << "Connect using SSO (Y): ";
        getline(wcin, tempStr);
        bSSOConnect = (tempStr == L"" || tempStr == L"Y" || tempStr == L"y");
        //if ( tempStr != L"" && tempStr != L"Y" && tempStr != L"y" )
        //    bSSOConnect = false;
    }

    // connect string examples
    //
    // single-hop user/pass:  Provider=MdrmSap.2;Data Source=Tableau's SAP Server A4H;SFC_CLIENT=800;SFC_LANGUAGE=;Prompt=NoPrompt;User ID=joeuser;Password=password;
    // single-hop SSO:        Provider=MdrmSap.2;Data Source=|SSO|Tableau's SAP Server A4H;SFC_CLIENT=800;SFC_LANGUAGE=;Prompt=NoPrompt;
    // double-hop via SST:    Provider=MDrmSap.2;Data Source=|SSO|Tableau's SAP Server A4H;SFC_CLIENT=800;SFC_LANGUAGE=;Prompt=NoPrompt;SFC_EXTIDTYPE=UN;SFC_EXTIDDATA=DEVELOPER3;

    wstringstream wss;
    wss << "Provider=" << provider << ";";
    wss << "Data Source=" << converter.from_bytes(connectionName) << ";";
    wss << "SFC_CLIENT=" << client << ";";
    wss << "SFC_LANGUAGE=" << language << ";";
    wss << "Prompt=NoPrompt;";

    string tempConnectString = "";
    if ( bImpersonateViaSST )
    {
        // impersonation via SST

        wcout << "Enter username to impersonate (" << username << "): ";
        getline(wcin, tempStr);
        if ( tempStr != L"" )
            username = tempStr;

        wss << "SFC_EXTIDTYPE=UN;";
        wss << "SFC_EXTIDDATA=" << username << ";";

        tempConnectString = trimNulls(wss.str());;
    }
    else if ( bSSOConnect )
    {
        // single-hop SSO

        tempConnectString = trimNulls(wss.str());;
    }
    else
    {
        // username password

        wcout << "Enter username (" << username << "): ";
        getline( wcin, tempStr );
        if ( tempStr != L"" )
            username = tempStr;

        string pw = getPassword();
        password = converter.from_bytes(pw);

        wss << "User ID=" << username << ";";

        // Don't include password in tempConnectString, which is included in output,
        // but do include it in wss, which is the connection string

        tempConnectString = trimNulls( wss.str() ) + "Password=********;";

        wss << "Password=" << password << ";";
    }

    connectionAttempt.client = trimNulls(client);
    connectionAttempt.language = trimNulls(language);
    if ( bImpersonateViaSST )
        connectionAttempt.authType = AuthType::SST;
    else
    {
        connectionAttempt.authType = bSSOConnect ? AuthType::SSO : AuthType::USERPASS;
    }
    connectionAttempt.connectString = tempConnectString;

    return wss.str();
}

void connect( const wstring& connStr, bool& bDataSourceOpen, bool& bSessionOpen, bool& bManagedCall )
{
    bool bImpersonate = false;
    initialize( bImpersonate );

    CDataSource ds;
    CSession session;
    hr = ds.OpenFromInitializationString( connStr.c_str() );
    if ( checkHResult( "OpenFromInitializationString", hr ) )
    {
        bDataSourceOpen = false;
    }

    if ( bDataSourceOpen )
    {
        session.Open( ds );
        if ( checkHResult( "Session Open", hr ) ) {
            ds.Close();
            bSessionOpen = false;
        }
    }

    /* Following block executes a query. Ideally there exists a standard query we can execute on any system

    auto managedCall = [&ds, &session, &bManagedCall]( char * desc, HRESULT hr ) {
        if (checkHResult(desc, hr)) {
            session.Close();
            ds.Close();
            bManagedCall = false;
        }
    };

    if (bDataSourceOpen && bSessionOpen)
    {
        wstring queryStr = L"SELECT\n{ [Measures].[0D_NW_DOCUM] } DIMENSION PROPERTIES[MEMBER_UNIQUE_NAME], [PARENT_UNIQUE_NAME], [PARENT_LEVEL], [MEMBER_CAPTION] ON COLUMNS\nFROM[$0D_NW_C01]";

        CCommand<CDynamicStringAccessor> recordset;
        managedCall("Open Records",
                    recordset.Open(session, queryStr.c_str()));
        int numColumns = recordset.GetColumnCount();
        for ( int col = 1; col <= numColumns; col++ ) {
            wcout << recordset.GetColumnName(col) << "\t|";
        }
        wcout << endl;
        do {
            for (int col = 1; col <= numColumns; col++) {
                wcout << recordset.GetString(col) << "\t|";
            }
            wcout << endl;
        } while (recordset.MoveNext() == S_OK);
    }
    */

    if ( bDataSourceOpen && bSessionOpen && bManagedCall )
    {
        session.Close();
        ds.Close();
    }

    outputConnectionAttempt();
}

void scanRfcTrace( bool bSSOConnect )
{
    string line;
    ifstream infile( "dev_rfc.trc" );
    if ( !infile.is_open() )
        return;

    bool bNoCredsSupplied = false;
    bool bUnableToEstablishSecurityContext = false;

    while ( getline( infile, line ) )
    {
        if ( line.find( "No credentials were supplied" ) != std::string::npos )
            bNoCredsSupplied = true;

        if ( line.find( "Unable to establish the security context" ) != std::string::npos )
            bUnableToEstablishSecurityContext = true;
    }
    infile.close();

    if ( bSSOConnect && bNoCredsSupplied && bUnableToEstablishSecurityContext )
        warningConditions.insert( SLC_NOT_LOGGED_IN );
}

// TODO error messages
// As currently coded, the file version will simply remain 0.0.0.0 if there
// is an error attempting to look it up. 
void fileVersion( string& path, WIN32_FIND_DATA& fileInfo, UINT version[4] )
{
    string pathFN = path + "\\" + converter.to_bytes( fileInfo.cFileName );
    CString pathFileName = (CString)pathFN.c_str();
    DWORD handle = 0;
    DWORD size = GetFileVersionInfoSizeW( pathFileName, &handle );

    if ( size == 0 )
        return;

    vector<BYTE> data( size );
    if ( !GetFileVersionInfo( pathFileName, 0, size, &data[0] ) )
        return;

    VS_FIXEDFILEINFO* pFileInfo = nullptr;
    UINT len = 0;
    if ( !VerQueryValue( &data[0], L"\\", (void**)&pFileInfo, &len ) )
        return;

    UINT major = HIWORD( pFileInfo->dwFileVersionMS );
    UINT minor = LOWORD( pFileInfo->dwFileVersionMS );

    UINT leastMajor = HIWORD( pFileInfo->dwFileVersionLS );
    UINT leastMinor = LOWORD( pFileInfo->dwFileVersionLS );

    version[0] = major;
    version[1] = minor;
    version[2] = leastMajor;
    version[3] = leastMinor;
}

void associatedFile( string& path, WIN32_FIND_DATA& fileInfo, dllStruct& dll )
{
    versionedFile asFile;
    asFile.location = path + "\\" + converter.to_bytes( fileInfo.cFileName );
    fileVersion( path, fileInfo, asFile.version );
    dll.associatedFiles.push_back( asFile );

    if ( asFile.location.compare(dll.file.location ) == 0 )
        dll.file.copyVersion( asFile.version );
}

void fileVersions( dllStruct& dll, string extension )
{
    const string& dllFileName = dll.file.location;

    WIN32_FIND_DATA fileInfo;
    HANDLE h;

    size_t idx = dllFileName.find_last_of( "\\" );
    if ( idx < 0 )
        return;

    string path = dllFileName.substr( 0, idx );
    CString pathTarget = (CString)path.c_str() + "\\*." + (CString)extension.c_str();

    h = FindFirstFile( pathTarget, &fileInfo );
    if ( h == INVALID_HANDLE_VALUE )
        return;

    associatedFile( path, fileInfo, dll);
    while ( FindNextFile( h, &fileInfo ) )
        associatedFile( path, fileInfo, dll);
}

//bool OSIs64Bit()
//{
//#if defined(BUILD_64BIT)
//    return true;
//#elif defined(BUILD_32BIT)
//    WINBOOL isWow64 = FALSE;
//
//    if ( IsWow64Process( GetCurrentProcess(), &isWow64 ) )
//        return isWow64 == TRUE;
//
//    return false;
//#endif
//}

void check32bitDriverVersion( const dllStruct& dll)
{
    if ( dll.file.compareVersion( MIN_DRIVER_VERSION_32_BIT ) < 0 )
        warningConditions.insert( BAD_32BIT_ODBO_PROVIDER_VERSION );
}

// TODO make this work when app is built for 32-bit OS? 
void odboProviders()
{
    bool b32found = false;
    bool b64found = false;

    HKEY key;
    REGSAM sam32 = KEY_READ | KEY_WOW64_32KEY;
    REGSAM sam64 = KEY_READ | KEY_WOW64_64KEY;

    // HKEY_CLASSES_ROOT is a shared reg key; its child nodes are visible
    // in the 32 and 64 bit registry views. This means we can't use the
    // progid to detect the bitness of the installed provider.
    // HKEY_CLASSES_ROOT\CLSID is redirected; i.e. not shared. So we use
    // the provider's CLSID to detect installed bitness.
    // LPCWSTR strKey = L"CLSID\\{B01952B0-AF66-11D1-B10D-0060086F6D97}"; // MDrmSap

    LPCWSTR strKey = L"CLSID\\{B01952B0-AF66-11D1-B10D-0060086F6D97}\\InprocServer32";

    LONG result = RegOpenKeyEx( HKEY_CLASSES_ROOT, strKey, 0, sam32, &key );

    if ( ERROR_SUCCESS == result )
    {
        b32found = true;

        wstring name = RegistryQueryValue( key, _T( "" ) );

        dllStruct dll;
        dll.bitness = "32";
        dll.file.location = trimNulls( name );
        fileVersions( dll, "dll" );
        vecOdboProviders.push_back(dll);

        check32bitDriverVersion(dll);

        RegCloseKey( key );
    }

    result = RegOpenKeyEx( HKEY_CLASSES_ROOT, strKey, 0, sam64, &key );

    if ( ERROR_SUCCESS == result )
    {
        b64found = true;

        wstring name = RegistryQueryValue( key, _T( "" ) );

        dllStruct dll;
        dll.bitness = "64";
        dll.file.location = trimNulls( name );
        fileVersions(dll, "dll" );
        vecOdboProviders.push_back(dll);

        RegCloseKey( key );
    }

    if ( !b32found )
        warningConditions.insert( NO_32BIT_ODBO_PROVIDER );
    if ( !b32found && !b64found )
        warningConditions.insert( NO_ODBO_PROVIDER );
    
    outputOdboProviders();
}

void checkSncLibContents( const dllStruct& dll, bool bImpersonateViaSST )
{
    if ( bImpersonateViaSST )
    {
        if ( dll.file.compareVersion( MIN_SNC_LIB_VERSION_SST) < 0 )
            warningConditions.insert(BAD_SNC_LIB_VERSION_SST);
    }
    else
    {
        if ( dll.file.compareVersion( MIN_SNC_LIB_VERSION_SSO) < 0 )
            warningConditions.insert(BAD_SNC_LIB_VERSION_SSO);
    }

    // check that op.file.location exists in op.associatedFiles
    for ( auto af : dll.associatedFiles )
    {
        if ( dll.file.location == af.location )
            return;
    }
    warningConditions.insert(MISSING_SNC_LIB_DLL);
}

// NOTE snclibContents considers !bImpersonateViaSST to 
//      imply bSSOConnect for simplicity
void snclibContents( bool bImpersonateViaSST )
{
    auto it = mapEnvironmentVariables.find("SNC_LIB");
    if ( it == mapEnvironmentVariables.end() )
        return; // TODO warn? 
    string snclib = it->second;

    // TODO consider how to handle bitness

    dllStruct dll;
    dll.bitness = "--";
    dll.file.location = snclib;
    fileVersions(dll, "dll");
    vecSnclibContents.push_back(dll);

    checkSncLibContents(dll, bImpersonateViaSST);

    outputSnclibContents();
}

void printDiffHeader()
{
    cout << "Diffs: " << endl;
    cout << "  Source     " << setw( 30 ) << "Connection Name" << "   " << setw( 3 ) << "Sncop" << "   " << setw(3) << "SncName" << endl;
}

void printDiff1( string type, definedConnectionStruct dc, bool& bDiffFound )
{
    cout << "  " << type << " only " << setw(32) << dc.name << "   " << setw(3) << dc.sncChoice << "     " << dc.sncName << endl;
    bDiffFound = true;
}

void printDiffBoth( definedConnectionStruct dcIni, definedConnectionStruct dcXML, bool& bDiffFound )
{
    cout << "  " << "Both     " << setw( 32 ) << dcIni.name << "   " << setw( 3 ) << dcIni.sncChoice << "     " << dcIni.sncName << endl;
    cout << "  " << "         " << setw( 32 ) << "" << "   " << setw( 3 ) << dcXML.sncChoice << "     " << dcXML.sncName << endl;
    bDiffFound = true;
}

void printDiffH(ConfigSource configSource)
{
    string path;
    string pathGlobal;
    string pathSource;
    string pathGlobalSource;

    configPathsAndSources(configSource, path, pathGlobal, pathSource, pathGlobalSource);
    cout << configSourceStrings[configSource] << endl;
    outputConfigPathsAndSources(configSource, path, pathGlobal, pathSource, pathGlobalSource);
}

void diffConnectionDefs()
{
    bool bDiffFound = false;

    environmentVariables( false );
    auto definedConnectionsIni = connectionDefsFromIni();

    printDiffH(ConfigSource::SAPLOGON_INI);

    locateLandscapeXMLFiles();
    auto definedConnectionsXML = connectionDefsFromXML();

    printDiffH(ConfigSource::LANDSCAPE_XML);

    auto itIni = definedConnectionsIni.begin();
    auto itXML = definedConnectionsXML.begin();

    printDiffHeader();
    while ( itIni != definedConnectionsIni.end() && itXML != definedConnectionsXML.end() )
    {
        auto dcIni = itIni->second;
        auto dcXML = itXML->second;

        if ( dcIni.name == dcXML.name )
        {
            if ( dcIni.sncChoice != dcXML.sncChoice || dcIni.sncName != dcXML.sncName )
            {
                printDiffBoth( dcIni, dcXML, bDiffFound );
            }
            itIni++;
            itXML++;
            continue;
        }
        else if ( dcIni.name < dcXML.name )
        {
            printDiff1( "INI", dcIni, bDiffFound );
            itIni++;
        }
        else 
        {
            printDiff1( "XML", dcXML, bDiffFound );
            itXML++;
        }
    }
    while ( itIni != definedConnectionsIni.end() )
    {
        auto dcIni = itIni->second;
        printDiff1( "INI", dcIni, bDiffFound );
        itIni++;
    }
    while ( itXML != definedConnectionsXML.end() )
    {
        auto dcXML = itXML->second;
        printDiff1( "XML", dcXML, bDiffFound );
        itXML++;
    }

    if ( !bDiffFound )
        cout << "  No diffs" << endl;
}

void usage()
{
    cout << "Usage: BWSSOTest [options]" << endl;
    cout << "-h, --help     show this help message and exit" << endl;
    cout << "-c <connName>  use specified connection name instead of selecting from defined connection list" << endl;
    cout << "-i <iniFile>   use specified full path to saplogon.ini instead of consulting SAPLOGON_INI_FILE environment variable" << endl;
    cout << "-s <i or l>    ignore LandscapeFormatEnabled, use saplogon.ini (i) or landscape xml (l) as source for connection definitions" << endl;
    cout << "-t <y or n>    use impersonation via server-side Trust" << endl;
    cout << "-d <y or n>    diff ini and xml connection definitions then exit" << endl;
}

void processArgs(int argc, char* argv[], bool& bSSOConnect, bool& bImpersonateViaSST, bool& bDiffConnectionDefs)
{
    cout << endl;

    if ( argc == 1 )
        return;

    if ( argc == 2 )
    {
        string key = argv[1];
        if ( key == "-h" || key == "--help" )
        {
            usage();
            exit(0);
        }
    }

    for ( int i = 1; i < argc - 1; i += 2 )
    {
        string key = argv[i];
        string val = argv[i + 1];
        if ( key == "-c" )
            connectionName = val;
        else if ( key == "-i" )
            sapLogonIniFileExplicit = converter.from_bytes(val);
        else if ( key == "-s" )
            connectionDefSourceExplicit = (val == "i" || val == "i") ? val : "";
        else if ( key == "-t" )
            bImpersonateViaSST = (val == "y" || val == "Y");
        else if ( key == "-d" )
            bDiffConnectionDefs = (val == "y" || val == "Y");
        else {
            cout << "unknown option '" << argv[i] << "'" << endl;
            usage();
            exit(1);
        }
    }

    if ( bImpersonateViaSST )
        bSSOConnect = false;
}

int main(int argc, char* argv[])
{
    bool bDiffConnectionDefs = false;
    bool bImpersonateViaSST = false;
    bool bSSOConnect = true;

    processArgs( argc, argv, bSSOConnect, bImpersonateViaSST, bDiffConnectionDefs );

    if ( bDiffConnectionDefs )
    {
        diffConnectionDefs();
        exit( 0 );
    }

    installedComponents();
    odboProviders();
    environmentVariables( bImpersonateViaSST );
    snclibContents( bImpersonateViaSST );
    sapLogonConfigSource();
    chooseConnection();

    bool bDataSourceOpen = true;
    bool bSessionOpen = true;
    bool bManagedCall = true;
    wstring connStr = getConnectString( bSSOConnect, bImpersonateViaSST );
    connect( connStr, bDataSourceOpen, bSessionOpen, bManagedCall );
    scanRfcTrace( bSSOConnect );

    cout << endl << "Test Complete" << endl << endl;

    outputConfigWarnings( bSSOConnect, bImpersonateViaSST );

    int exitCode = (bDataSourceOpen && bSessionOpen && bManagedCall) ? 0 : 1;
    outputExitCode(exitCode);

    persistResults();

    system( "PAUSE" );

    exit( exitCode );
}

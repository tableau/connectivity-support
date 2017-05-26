# connectivity-support
Provides a set of tools for diagnosing common infrastructure concerns in the connectivity layer.

## SAPBW

### BWSSOTestTool
This is a Windows console application that looks up SAP BW-related configuration information and attempts to connect to a specified BW instance, producing output that can help diagnose configuration issues. It captures relevant configuration settings, gives warnings when it finds settings that are missing or different than expected, and produces trace files for the connection attempt. It is useful for both kerberos-based (think Tableau Desktop to BW) and delegation via Server Side Trust (think Viewer Credentials authentication in Tableau Server). 

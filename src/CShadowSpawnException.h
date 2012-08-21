/* 
Copyright (c) 2011 Craig Andera (shadowspawn@wangdera.com)

Permission is hereby granted, free of charge, to any person obtaining
a copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:

The above copyright notice and this permission notice shall be
included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
*/

#pragma once

class CShadowSpawnException
{
private:
    CString _message;
	HRESULT _result;
public: 
    CShadowSpawnException::CShadowSpawnException(HRESULT result, const LPCTSTR message)
    {
        _message.Append(message);
		_result = result;
    }
	CShadowSpawnException::CShadowSpawnException(DWORD result, const LPCTSTR message)
    {
        _message.Append(message);
		_result = HRESULT_FROM_WIN32(result);
    }
	CShadowSpawnException::CShadowSpawnException(const LPCTSTR message)
    {
        _message.Append(message);
		_result = HRESULT_FROM_WIN32(::GetLastError());
    }
    LPCTSTR get_Message(void)
    {
        return _message; 
    }
	 HRESULT get_HResult(void)
    {
        return _result; 
    }
};
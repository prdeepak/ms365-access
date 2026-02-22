#!/bin/bash
echo "CWD: $(pwd)" >> /tmp/mcp-ms365-stderr.log
echo "ARGS: $@" >> /tmp/mcp-ms365-stderr.log
cd /Users/deepakramachandran/bin/ms365-access/backend
export PYTHONPATH="/Users/deepakramachandran/marvin:$PYTHONPATH"
exec /Users/deepakramachandran/bin/ms365-access/backend/.venv/bin/python -m mcp_server 2>>/tmp/mcp-ms365-stderr.log

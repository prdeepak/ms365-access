#!/bin/bash
echo "CWD: $(pwd)" >> /tmp/mcp-ms365-stderr.log
echo "ARGS: $@" >> /tmp/mcp-ms365-stderr.log
cd /Users/deepak/bin/ms365-access/backend
exec /Users/deepak/bin/ms365-access/backend/venv/bin/python -m mcp_server 2>>/tmp/mcp-ms365-stderr.log

### How do I build the duplicate detection exe for Windows envs?

```
pyinstaller --onefile nedss_duplicate_finder.py
```

The above command will generate `nedss_duplicate_finder.exe` which is a stand-alone executable ready for deployment to Windows envs.

When `nedss_duplicate_finder.exe` is launched, the program will run in interactive mode (prompting the user for required params).

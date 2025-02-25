This tool provides fixes and access to multiple applications by running

```irm overbytestech.com/repair | iex```

Many servers and a few workstations do not like running this script as it does not use SSL, or scripts are not allowed.
To get around this, you can temporarily change the execution policy (for the current session only) by using this command

`Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process`

Otherwise you can do `Set-ExecutionPolicy RemoteSigned` to change the policy on the entire system.

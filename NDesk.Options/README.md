# NDesk.Options.Fork

### Build Status
[![Build status](https://ci.appveyor.com/api/projects/status/9rkhbjxrka6ra66a/branch/master?svg=true)](https://ci.appveyor.com/project/torston/ndesk-options-fork/branch/master) [![Test status](http://flauschig.ch/batch.php?type=tests&account=torston&slug=ndesk-options-fork)](https://ci.appveyor.com/api/projects/status/9rkhbjxrka6ra66a/branch/master) [![NuGet version](https://badge.fury.io/nu/Ndesk.Options.Fork.svg)](https://badge.fury.io/nu/Ndesk.Options.Fork)
## Description
This repository is the fork of NDesk.Options 0.2.1 (callback-based program option parser for C#).

Original project link: http://www.ndesk.org/Options

Original project documentation: http://www.ndesk.org/doc/ndesk-options/NDesk.Options/
## Quickstart
1) Create Option Set
```c#
var p = new OptionSet ()
```
2) Add argumets ( Important: if you need to get value in lambda you need `=` like: `g|game=`)
```c#
p.Add("s|status=", n => Console.WriteLine("Status is "+ n));
```
3) If it find your argument the lambda will called, othewise it will return it back
```c#
var unexpectedArguments = p.Parse(argsArray);

foreach(var arg in unexpectedArguments)
{
    Console.WriteLine($"Unknown argument: {arg}");
}
```
 
Output:
```
> program.exe -s Ready
Status is Ready
> program.exe --status "Loading"
Status is Loading
> program.exe --anotherArgument 12
Unknown argument: --anotherArgument
Unknown argument: 12
```
## Getting Deeper 
### Define options
```c#
 var p = new OptionSet ()
 
 // You can call: -n "Rick" or --name "Morty"
 p.Add("n|name=", n => Console.WriteLine(n));
 
 // You can call only with long argument: --name "Morty"
 p.Add("name=", n => Console.WriteLine("First Name: " + n));
  
  // Bool options usage: -s, you dont need `=` in case of bool option
  p.Add("s|isSmart", s => Console.WriteLine(s != null));
  
  // Int options usage: -a 11
  p.Add("a|age=", (int a) => Console.WriteLine("Age: " + s));
 ```
### Parce options
```c#
var unexpectedArguments = p.Parse (args);

foreach(var arg in unexpectedArguments)
{
    Console.WriteLine($"Unknown argument: {arg}");
}
 ```
 
#### Command Line: 
```
program.exe --name "Morty" --surname "Smith" --sex "male" -a 13
 
First Name: Morty
Last Name: Smith
Age: 13
Unknown argument: --sex
Unknown argument: male
```
#### License

The project is licensed under the [MIT](https://opensource.org/licenses/mit-license.php) license.

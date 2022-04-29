// See https://aka.ms/new-console-template for more information
//Console.WriteLine("Hello, World!");

using case1;

var solution = new Solution(Path.Combine(Environment.CurrentDirectory, @".\Case1.xlsx"));

solution.sortSheetThreeColumnAscending("SCont_done");

solution.saveAndClose();
using RazorEngine;
using RazorEngine.Templating;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CommandTools.Templates
{
    static class Render
    {
        private const string keyTemplate = "_keyTemplate";
        public static string Execute(string filePahtTemplate, object obj)
        {
            var template = File.OpenText(filePahtTemplate).ReadToEnd();
            return Engine.Razor.RunCompile( template, keyTemplate, null, obj);
        }
    }
}

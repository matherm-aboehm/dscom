// Copyright 2022 dSPACE GmbH, Mark Lechtermann, Matthias Nissen and Contributors
// 
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
// 
//     http://www.apache.org/licenses/LICENSE-2.0
// 
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

using System.Reflection;

namespace dSPACE.Runtime.InteropServices.Writer;

internal sealed class PropertySetMethodWriter : PropertyMethodWriter
{
    internal new readonly record struct FactoryArgs(InterfaceWriter InterfaceWriter, PropertyInfo PropertyInfo, MethodInfo MethodInfo, WriterContext Context, string MethodName)
        : WriterFactory.IWriterArgsFor<PropertySetMethodWriter>
    {
        PropertySetMethodWriter WriterFactory.IWriterArgsFor<PropertySetMethodWriter>.CreateInstance()
        {
            return new PropertySetMethodWriter(InterfaceWriter, PropertyInfo, MethodInfo, Context, MethodName);
        }
    }

    private PropertySetMethodWriter(InterfaceWriter interfaceWriter, PropertyInfo propertyInfo, MethodInfo methodInfo, WriterContext context, string methodName) : base(interfaceWriter, propertyInfo, methodInfo, context, methodName)
    {
        InvokeKind = INVOKEKIND.INVOKE_PROPERTYPUT;
    }

    public override void Create()
    {
        try
        {
            if (MethodInfo.GetParameters().Any(p =>
             {
                 return p.ParameterType.Equals(typeof(object)) ||
                 p.ParameterType.ToString() == typeof(IDispatch).FullName ||
                     p.ParameterType.IsInterface;
             }))
            {
                InvokeKind = INVOKEKIND.INVOKE_PROPERTYPUTREF;
            }
        }
        catch (FileNotFoundException)
        {
            //FileNotFoundException can occur on GetParameters() method, disable method writer
        }

        base.Create();
    }

    protected override List<string> GetNamesForParameters()
    {
        var names = base.GetNamesForParameters();

        if (names.Count > 1 && string.Equals(names[names.Count - 1], "value", StringComparison.OrdinalIgnoreCase))
        {
            names = names.GetRange(0, names.Count - 1);
        }

        return names;
    }

    protected override short GetParametersCount()
    {
        return (short)MethodInfo.GetParameters().Length;
    }
}

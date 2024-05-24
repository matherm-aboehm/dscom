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

internal sealed class PropertyGetMethodWriter : PropertyMethodWriter
{
    internal new readonly record struct FactoryArgs(InterfaceWriter InterfaceWriter, PropertyInfo? PropertyInfo, MethodInfo MethodInfo, WriterContext Context, string MethodName)
        : WriterFactory.IWriterArgsFor<PropertyGetMethodWriter>
    {
        PropertyGetMethodWriter WriterFactory.IWriterArgsFor<PropertyGetMethodWriter>.CreateInstance()
        {
            return new PropertyGetMethodWriter(InterfaceWriter, PropertyInfo, MethodInfo, Context, MethodName);
        }
    }

    private PropertyGetMethodWriter(InterfaceWriter interfaceWriter, PropertyInfo? propertyInfo, MethodInfo methodInfo, WriterContext context, string methodName) : base(interfaceWriter, propertyInfo, methodInfo, context, methodName)
    {
        InvokeKind = INVOKEKIND.INVOKE_PROPERTYGET;
    }

    protected override short GetParametersCount()
    {
        // If case of HRESULT as return value, the parameter count should be +1
        return (short)(UseHResultAsReturnValue ? MethodInfo.GetParameters().Length + 1 : MethodInfo.GetParameters().Length);
    }
}

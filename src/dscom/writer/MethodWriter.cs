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

#if NET5_0_OR_GREATER
using System.Diagnostics.CodeAnalysis;
#endif
using System.Reflection;
using System.Runtime.InteropServices;
using FUNCFLAGS = System.Runtime.InteropServices.ComTypes.FUNCFLAGS;

namespace dSPACE.Runtime.InteropServices.Writer;

internal class MethodWriter : BaseWriter, WriterFactory.IProvidesFinishCreateInstance
{
    internal readonly record struct FactoryArgs(InterfaceWriter InterfaceWriter, MethodInfo MethodInfo, WriterContext Context, string MethodName)
        : WriterFactory.IWriterArgsFor<MethodWriter>
    {
        MethodWriter WriterFactory.IWriterArgsFor<MethodWriter>.CreateInstance()
        {
            return new MethodWriter(InterfaceWriter, MethodInfo, Context, MethodName);
        }
    }

    protected MethodWriter(InterfaceWriter interfaceWriter, MethodInfo methodInfo, WriterContext context, string methodName) : base(context)
    {
        InterfaceWriter = interfaceWriter;
        MethodInfo = methodInfo;
        MethodName = methodName;

        //switch off HResult transformation on method level
        PreserveSig = (methodInfo.MethodImplementationFlags & MethodImplAttributes.PreserveSig) != 0
                    || methodInfo.GetCustomAttribute<PreserveSigAttribute>() != null;
        UseHResultAsReturnValue = !PreserveSig && interfaceWriter.UseHResultAsReturnValue;
    }

#if NET5_0_OR_GREATER
    [MemberNotNull(nameof(MemberInfo))]
#endif
    protected virtual void FinishCreateInstance()
    {
        //HINT: The following type of initialization logic only works here, not in base ctor.
        // In case of a Property
        MemberInfo ??= MethodInfo;
        // So derived ctors can set MemberInfo (or something similiar) and everything else
        // which is depending on it will then be initialized here after ctor completed.

        CreateParameterWriters();
        // So as a conclusion, this or a derived class can only be partially constructed by
        // its ctor and FinishCreateInstance() needs to be called immediately after ctor to
        // finish the construction.
    }

    void WriterFactory.IProvidesFinishCreateInstance.FinishCreateInstance()
    {
        FinishCreateInstance();
    }

    public bool PreserveSig { get; }

    public bool UseHResultAsReturnValue { get; private set; }

    public InterfaceWriter InterfaceWriter { get; }

    //HINT: Cannot use NotNullAttribute in combination with nullable-ref type here,
    // because older analyzers ignore it and still show warnings for null references.
    // So make it non-null and initialize it to null with null! to remove the warning for
    // constructors.
    // Make sure to always call FinishCreateInstance after construction, so that it
    // actually has a value after that and not just trick the analyzer.
    //[NotNull]
    public MemberInfo MemberInfo { get; protected set; } = null!;

    internal MethodInfo MethodInfo { get; }

    internal string MethodName { get; private set; }

    protected INVOKEKIND InvokeKind { get; set; } = INVOKEKIND.INVOKE_FUNC;

    private List<ParameterWriter> ParameterWriters { get; } = new();

    private ParameterWriter? ReturnParamWriter { get; set; }

    protected virtual short GetParametersCount()
    {
        var retVal = (short)(UseHResultAsReturnValue && !MethodInfo.ReturnType.Equals(typeof(void)) ?
                MethodInfo.GetParameters().Length + 1 :
                MethodInfo.GetParameters().Length);
        return retVal;
    }

    private bool? _isComVisible;

    protected virtual bool IsComVisible
    {
        get
        {
            if (_isComVisible != null)
            {
                return _isComVisible.Value;
            }

            var methodAttribute = MethodInfo.GetCustomAttribute<ComVisibleAttribute>();
            _isComVisible = methodAttribute == null || methodAttribute.Value;
            return _isComVisible.Value;
        }
    }

    /// <summary>
    /// Gets a value indicating whether this instance of MethodWriter is valid to generate a FuncDesc.
    /// </summary>
    /// <value>true if valid; otherwise false</value>
    public bool IsValid => IsVisibleMethod && HasValidParameters;

    public bool HasValidParameters => ParameterWriters.All(z => z.IsValid) && ReturnParamWriter != null && ReturnParamWriter.IsValid;

    public bool IsVisibleMethod => (MethodInfo != MemberInfo || !MethodInfo.IsGenericMethod) && IsComVisible;

    public int FunctionIndex { get; set; } = -1;

    public int VTableOffset { get; set; }

    public void CreateParameterWriters()
    {
        //check for disabled HRESULT transformation
        var useNetMethodSignature = !PreserveSig && !UseHResultAsReturnValue;
        foreach (var paramInfo in MethodInfo.GetParameters())
        {
            ParameterWriters.Add(new ParameterWriter(this, paramInfo, Context, false));
        }
        if (useNetMethodSignature || PreserveSig)
        {
            var returnType = MethodInfo.ReturnType;
            var returnParam = MethodInfo.ReturnParameter;
            // if it is a custom value type that is namend HRESULT then
            // don't use that for the ReturnParamWriter, instead enable
            // the HRESULT transformation
            if (returnType.IsValueType && returnType.Name == "HRESULT")
            {
                returnParam = new HResultParamInfo();
            }
            ReturnParamWriter = new ParameterWriter(this, returnParam, Context, false);
        }
        else
        {
            // do not create retval parameter for void return
            if (!MethodInfo.ReturnType.Equals(typeof(void)))
            {
                ParameterWriters.Add(new ParameterWriter(this, MethodInfo.ReturnParameter, Context, true));
            }
            ReturnParamWriter = new ParameterWriter(this, new HResultParamInfo(), Context, false);
        }
    }

    public override void Create()
    {
        ReturnParamWriter!.Create();

        foreach (var paramWriter in ParameterWriters)
        {
            paramWriter.Create();
        }

        var vTableOffset = VTableOffset;
        var funcIndex = FunctionIndex;
        if (funcIndex == -1)
        {
            throw new InvalidOperationException("Function index is -1");
        }

        var typeInfo = InterfaceWriter.TypeInfo;

        if (!IsComVisible)
        {
            return;
        }

        if (ReturnParamWriter != null)
        {
            foreach (var writer in ParameterWriters.Append(ReturnParamWriter))
            {
                writer.ReportEvent();
            }
        }

        if (!IsValid)
        {
            return;
        }

        if (typeInfo == null)
        {
            throw new ArgumentException("ICreateTypeInfo2 is null.");
        }

        var memidCreated = (int)InterfaceWriter.DispatchIdCreator!.GetDispatchId(this);
        string[] names;
        try
        {
            FUNCFLAGS flags = 0;
            var flagsAttrs = MethodInfo.GetCustomAttributes<Attributes.FuncFlagsAttribute>();
            if (flagsAttrs != null && flagsAttrs.Any())
            {
                foreach (var flagAttr in flagsAttrs)
                {
                    flags = flagAttr.UpdateFlags(flags);
                }
            }

            FUNCDESC? funcDesc = new FUNCDESC
            {
                callconv = CALLCONV.CC_STDCALL,
                cParams = GetParametersCount(),
                cParamsOpt = 0,
                cScodes = 0,
                elemdescFunc = ReturnParamWriter!.ElementDescription,
                funckind = InterfaceWriter.FuncKind,
                invkind = (memidCreated == 0 && InvokeKind == INVOKEKIND.INVOKE_FUNC) ? INVOKEKIND.INVOKE_PROPERTYGET : InvokeKind,
                lprgelemdescParam = GetElementDescriptionsPtrForParameters(),
                lprgscode = IntPtr.Zero,
                memid = memidCreated,
                oVft = (short)vTableOffset,
                wFuncFlags = (short)flags
            };

            //Check if method still enabled. If a parameter is not enabled, the method should not be created.
            if (!IsValid)
            {
                return;
            }

            names = GetNamesForParameters().ToArray();

            typeInfo.AddFuncDesc((uint)funcIndex, funcDesc.Value)
                .ThrowIfFailed($"Failed to add function description for {MethodInfo.DeclaringType}.{MethodInfo.Name}().");

            typeInfo.SetFuncAndParamNames((uint)funcIndex, names, (uint)names.Length)
                .ThrowIfFailed($"Failed to set function and parameter names for {MethodInfo.DeclaringType}.{MethodInfo.Name}().");

            var description = MethodInfo.GetCustomAttribute<System.ComponentModel.DescriptionAttribute>();
            if (description != null)
            {
                typeInfo.SetFuncDocString((uint)funcIndex, description.Description)
                    .ThrowIfFailed($"Failed to set function documentation string for {MethodInfo.DeclaringType}.{MethodInfo.Name}().");
            }

            if (memidCreated == 0 && InvokeKind == INVOKEKIND.INVOKE_FUNC)
            {
                //Function has been forced as property (default member handling)
                typeInfo.SetFuncCustData((uint)funcIndex, new Guid(Guids.GUID_Function2Getter), 1)
                    .ThrowIfFailed($"Failed to set function custom data for {MethodInfo.DeclaringType}.{MethodInfo.Name}().");
            }

        }
        catch (Exception ex) when (ex is TypeLoadException or COMException)
        {
            Context.LogWarning(ex.Message, HRESULT.E_INVALIDARG);
        }
    }

    protected virtual List<string> GetNamesForParameters()
    {
        var names = new List<string>();
        var methodName = Context.NameResolver.GetMappedName(MethodInfo, MethodName);
        names.Add(methodName);

        MethodInfo.GetParameters().ToList().ForEach(p => names.Add(Context.NameResolver.GetMappedName(p, p.Name ?? string.Empty) ?? string.Empty));

        if (UseHResultAsReturnValue && !MethodInfo.ReturnType.Equals(typeof(void)))
        {
            names.Add("pRetVal");
        }

        return names;
    }

    private IntPtr GetElementDescriptionsPtrForParameters()
    {
        var parameters = new ELEMDESC[ParameterWriters.Count];

        foreach (var paramWriter in ParameterWriters)
        {
            parameters[ParameterWriters.IndexOf(paramWriter)] = paramWriter.ElementDescription;
        }

        var intPtr = StructuresToPtr(parameters);

        return intPtr;
    }
}

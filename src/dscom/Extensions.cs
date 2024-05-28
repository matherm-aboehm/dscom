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

using System.Collections.Immutable;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Reflection.Metadata;
using System.Reflection.PortableExecutable;
using System.Runtime.ConstrainedExecution;
using System.Runtime.InteropServices;

namespace dSPACE.Runtime.InteropServices;

[ExcludeFromCodeCoverage]
internal static class Extensions
{
#if NETFRAMEWORK || NETSTANDARD
    /// <summary>
    /// Deconstructs the current <see cref="KeyValuePair&lt;TKey, TValue&gt;"/>.
    /// </summary>
    /// <typeparam name="TKey">The type of the key.</typeparam>
    /// <typeparam name="TValue">The type of the value.</typeparam>
    /// <param name="pair">The current <see cref="KeyValuePair&lt;TKey, TValue&gt;"/> to deconstruct.</param>
    /// <param name="key">The key of the current <see cref="KeyValuePair&lt;TKey, TValue&gt;"/>.</param>
    /// <param name="value">The value of the current <see cref="KeyValuePair&lt;TKey, TValue&gt;"/>.</param>
    internal static void Deconstruct<TKey, TValue>(this KeyValuePair<TKey, TValue> pair, out TKey key, out TValue value)
    {
        key = pair.Key;
        value = pair.Value;
    }
#endif

    [SuppressMessage("Microsoft.Style", "IDE0060", Justification = "For future use")]
    internal static int GetTLBLanguageIdentifierForAssembly(this Assembly assembly)
    {
        // This is currently a limitation.
        // Only LCID_NEUTRAL is supported.
        return Constants.LCID_NEUTRAL;
    }

    internal static Version GetTLBVersionForAssembly(this Assembly assembly)
    {
        TypeLibVersionAttribute? typeLibVersionAttribute = null;
        try
        {
            typeLibVersionAttribute = assembly.GetCustomAttribute<TypeLibVersionAttribute>();
        }
        catch (FileNotFoundException)
        {
        }

        if (typeLibVersionAttribute != null)
        {
            return new Version(typeLibVersionAttribute.MajorVersion, typeLibVersionAttribute.MinorVersion);
        }
        if (assembly.GetName().Version != null)
        {
            return assembly.GetName().Version!;
        }

        throw new InvalidOperationException();
    }

    internal static Guid GetTLBGuidForAssembly(this Assembly assembly)
    {
        return MarshalExtension.GetTypeLibGuidForAssembly(assembly);
    }

    internal static bool IsSpecialHandledClass(this Type type)
    {
        var typeAsString = type.ToString();
        return typeAsString switch
        {
            "System.Object" => true,
            "System.String" => true,
            _ => false
        };
    }

    /// <summary>
    /// Check if the type is System.Void, System.Drawing, System.DateTime, System.Decimal or System.GUID.
    /// </summary>
    /// <param name="type">The type to check.</param>
    /// <returns>Return true if the type is a special type; otherwise false.</returns>
    internal static bool IsSpecialHandledValueType(this Type type)
    {
        var typeAsString = type.ToString();
        return typeAsString switch
        {
            "System.Void" => true,
            "System.Drawing.Color" => true,
            "System.DateTime" => true,
            "System.Decimal" => true,
            "System.Guid" => true,
            _ => false
        };
    }

    internal static Type GetCLRType(this VarEnum varenum)
    {
        return varenum switch
        {
            VarEnum.VT_VOID => typeof(void),
            VarEnum.VT_DECIMAL => typeof(decimal),
            VarEnum.VT_R4 => typeof(float),
            VarEnum.VT_R8 => typeof(double),
            VarEnum.VT_BOOL => typeof(bool),
            VarEnum.VT_VARIANT => typeof(object),
            VarEnum.VT_BSTR => typeof(string),
            VarEnum.VT_I4 => typeof(int),
            VarEnum.VT_UI1 => typeof(byte),
            VarEnum.VT_I1 => typeof(sbyte),
            VarEnum.VT_I2 => typeof(short),
            VarEnum.VT_I8 => typeof(long),
            VarEnum.VT_UI2 => typeof(ushort),
            VarEnum.VT_UI4 => typeof(uint),
            VarEnum.VT_UI8 => typeof(ulong),
            VarEnum.VT_SAFEARRAY => typeof(object[]),
            VarEnum.VT_DATE => typeof(DateTime),
            VarEnum.VT_UNKNOWN => typeof(IUnknown),
            VarEnum.VT_DISPATCH => typeof(IDispatch),
            VarEnum.VT_PTR => typeof(IntPtr),
            _ => throw new NotSupportedException(),
        };
    }

    internal static UnmanagedType? ToUnmanagedType(this VarEnum varenum)
    {
        return varenum switch
        {
            VarEnum.VT_EMPTY => null,
            VarEnum.VT_NULL => null,
            VarEnum.VT_I2 => UnmanagedType.I2,
            VarEnum.VT_I4 => UnmanagedType.I4,
            VarEnum.VT_R4 => UnmanagedType.R4,
            VarEnum.VT_R8 => UnmanagedType.R8,
            VarEnum.VT_CY => UnmanagedType.Bool,
            VarEnum.VT_DATE => null,
            VarEnum.VT_BSTR => UnmanagedType.BStr,
            VarEnum.VT_DISPATCH => UnmanagedType.IDispatch,
            VarEnum.VT_ERROR => UnmanagedType.Error,
            VarEnum.VT_BOOL => UnmanagedType.Bool,
            VarEnum.VT_VARIANT => UnmanagedType.Struct,
            VarEnum.VT_UNKNOWN => UnmanagedType.IUnknown,
            VarEnum.VT_DECIMAL => UnmanagedType.Currency,
            VarEnum.VT_I1 => UnmanagedType.I1,
            VarEnum.VT_UI1 => UnmanagedType.U1,
            VarEnum.VT_UI2 => UnmanagedType.U2,
            VarEnum.VT_UI4 => UnmanagedType.U4,
            VarEnum.VT_I8 => UnmanagedType.I8,
            VarEnum.VT_UI8 => UnmanagedType.U8,
            VarEnum.VT_INT => UnmanagedType.I4,
            VarEnum.VT_UINT => UnmanagedType.U4,
            VarEnum.VT_VOID => null,
            VarEnum.VT_HRESULT => UnmanagedType.Error,
            VarEnum.VT_PTR => null,
            VarEnum.VT_SAFEARRAY => UnmanagedType.SafeArray,
            VarEnum.VT_CARRAY => null,
            VarEnum.VT_USERDEFINED => null,
            VarEnum.VT_LPSTR => UnmanagedType.LPStr,
            VarEnum.VT_LPWSTR => UnmanagedType.LPWStr,
            VarEnum.VT_RECORD => null,
            VarEnum.VT_FILETIME => null,
            VarEnum.VT_BLOB => null,
            VarEnum.VT_STREAM => null,
            VarEnum.VT_STORAGE => null,
            VarEnum.VT_STREAMED_OBJECT => null,
            VarEnum.VT_STORED_OBJECT => null,
            VarEnum.VT_BLOB_OBJECT => null,
            VarEnum.VT_CF => null,
            VarEnum.VT_CLSID => null,
            VarEnum.VT_VECTOR => null,
            VarEnum.VT_ARRAY => null,
            VarEnum.VT_BYREF => null,
            _ => throw new NotSupportedException(),
        };
    }

    internal static bool IsComVisible(this Type type)
    {
        type = type.GetUnderlayingType();
        var AssemblyIsComVisible = type.Assembly.GetCustomAttribute<ComVisibleAttribute>()?.Value ?? true;

        if (AssemblyIsComVisible)
        {
            //a type of a com visible assembly is com visible unless it is explicitly com invisible
            return type.GetCustomAttribute<ComVisibleAttribute>()?.Value ?? true;
        }
        else
        {
            //a type of a com invisible assembly is com invisible unless it is explicitly com visible
            return type.GetCustomAttribute<ComVisibleAttribute>()?.Value ?? false;
        }
    }

    internal static Type GetUnderlayingType(this Type type)
    {
        var returnType = type;
        while (returnType!.IsByRef || returnType!.IsArray)
        {
            returnType = returnType.GetElementType();
        }
        return returnType;
    }

    internal static ClassInterfaceType GetComClassInterfaceType(this Type type)
    {
        if (type.IsGenericType)
        {
            return ClassInterfaceType.None;
        }
        var baseType = type.BaseType;
        while (baseType != null)
        {
            if (baseType.IsGenericType)
            {
                return ClassInterfaceType.None;
            }
            baseType = baseType.BaseType;
        }

        //skip ClassSupportsIClassX check, it is not implemented yet

        var defaultClassInterfaceType = type.Assembly.GetCustomAttribute<ClassInterfaceAttribute>()?.Value ?? ClassInterfaceType.AutoDispatch;
        return type.GetCustomAttribute<ClassInterfaceAttribute>()?.Value ?? defaultClassInterfaceType;
    }

    internal static bool IsDelegate(this Type type)
    {
        return type.BaseType?.Equals(typeof(MulticastDelegate)) ?? false;
    }

    internal static bool IsSharedByGenericInstantiations(this Type type)
    {
        return type.GetGenericArguments().Where(t => !t.IsGenericParameter &&
            (t.FullName == "System.__Canon" ||
            (t.IsGenericType && t.IsSharedByGenericInstantiations()))).Any();
    }

    internal static bool IsNestedPublicRecursive(this Type type)
    {
        while (type.IsNestedPublic)
        {
            type = type.DeclaringType!;
        }
        return type.IsPublic;
    }

    internal static IEnumerable<T> GetCustomAttributesRecursive<T>(this Type element) where T : Attribute
    {
        var result = element.GetCustomAttributes<T>();
        if (element.BaseType != null)
        {
            result = result.Union(element.BaseType.GetCustomAttributesRecursive<T>());
        }
        return result;
    }

    internal static List<List<Type>> GetComInterfacesRecursive(this Type type)
    {
        var result = new List<List<Type>>
        {
            type.GetInterfaces().Where(x => x.IsComVisible()).ToList()
        };
        if (type.BaseType != null)
        {
            var innerResult = type.BaseType.GetComInterfacesRecursive();
            result.AddRange(innerResult);
        }
        if (result.First() != null)
        {
            var allOtherTypes = result.Skip(1).SelectMany(static y => y);
            result[0] = result[0].Except(allOtherTypes).ToList();
        }
        return result;
    }

    internal static Type? GetComPlusParentType(this Type type)
    {
        var parentType = type.BaseType;
        if (parentType != null && parentType.IsImport)
        {
            if (parentType.IsProjectedFromWinRT())
            {
                // skip all Com Import classes
                do
                {
                    parentType = parentType.BaseType;
                    Debug.Assert(parentType != null);
                } while (parentType!.IsImport);

                // Now we have either System.__ComObject or WindowsRuntime.RuntimeClass
                if (parentType.FullName != "System.__ComObject")
                {
                    return parentType;
                }
            }
            else
            {
                // Skip the single ComImport class we expect
                Debug.Assert(parentType.BaseType != null);
                parentType = parentType.BaseType;
            }
            Debug.Assert(!parentType!.IsImport);

            // Skip over System.__ComObject, expect System.MarshalByRefObject
            parentType = parentType.BaseType;
            Debug.Assert(parentType != null);
            Debug.Assert(parentType!.IsMarshalByRef);
            Debug.Assert(parentType.BaseType != null);
            Debug.Assert(parentType.BaseType!.Equals(typeof(object)));
        }
        return parentType;
    }

    internal static TDelegate CreateDelegate<TDelegate>(this MethodInfo mi, object? obj)
        where TDelegate : Delegate
    {
        return (TDelegate)mi.CreateDelegate(typeof(TDelegate), obj);
    }

    private class BlobReaderContext : CriticalFinalizerObject, IDisposable
    {
        private readonly PEReader _peReader;
        private int _disposedValue;

        public BlobReaderContext(PEReader peReader)
        {
            _peReader = peReader;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (Interlocked.Exchange(ref _disposedValue, 1) == 0)
            {
                _peReader.Dispose();
            }
        }

        ~BlobReaderContext()
        {
            Dispose(disposing: false);
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }

    internal static BlobReader GetBlobReaderFromByteArray(byte[] array, out IDisposable? blobReaderContext)
    {
        //HINT: PEReader doesn't read anything here, it is only used to convert
        //byte array to a BlobReader with the help of public APIs instead of
        //reflection over internals or using unsafe code.
        PEReader peReader = new(array.ToImmutableArray());
        blobReaderContext = new BlobReaderContext(peReader);
        IDisposable? disposeOnException = null;
        try
        {
            var block = peReader.GetEntireImage();
            return block.GetReader();
        }
        catch
        {
            (disposeOnException, blobReaderContext) =
                (blobReaderContext, disposeOnException);
            throw;
        }
        finally
        {
            disposeOnException?.Dispose();
        }
    }
}

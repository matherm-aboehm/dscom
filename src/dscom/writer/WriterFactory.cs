using System.Reflection;

namespace dSPACE.Runtime.InteropServices.Writer;

internal static class WriterFactory
{
    internal interface IWriterArgsFor<TWriter> where TWriter : BaseWriter
    {
        TWriter CreateInstance();
    }

    internal interface IProvidesFinishCreateInstance
    {
        void FinishCreateInstance();
    }

    /// <summary>
    /// Static factory method for creating an instance of <typeparamref name="TWriter"/> type.
    /// This supports the factory pattern to give the type the chance to have a different
    /// complex construction logic, which can not be solved with default C# constructors.
    /// <see cref="IProvidesFinishCreateInstance"/> has to be implemented on the type, when
    /// the constructor itself is only one half of the construction logic and the type needs
    /// to make sure something else is done immediately after the constructor has returned.
    /// </summary>
    /// <example>
    /// For example (<see cref="InterfaceWriter"/>) when base ctor needs to call virtual
    /// methods which will be overriden in derived types:
    /// abstract class Foo
    /// {
    ///     protected Foo() => Bar(); //CA2214
    ///     protected abstract void Bar();
    /// }
    /// class FooBar : Foo
    /// {
    ///     public FooBar() : base()
    ///     {
    ///         Msg = "Hello World!"; //to late to initialize here because Bar should be called from base class
    ///     }
    ///     public string Msg { get; }
    ///     protected override void Bar() => Console.WriteLine($"FooBar.Bar() called with: {Msg}.");
    /// }
    /// This factory pattern solves <see href="https://learn.microsoft.com/en-us/dotnet/fundamentals/code-analysis/quality-rules/ca2214">CA2214</see>
    /// and the initialization problem from above example.
    /// See <see cref="IProvidesFinishCreateInstance"/> implementation on <see cref="MethodWriter"/> for another example.
    /// </example>
    /// <remarks>Someday when Type Classes made it into C# this would be canditate to refactor.
    /// (<see href="https://github.com/dotnet/csharplang/issues/110"/>)</remarks>
    /// <typeparam name="TWriter">The writer type that this factory creates.</typeparam>
    /// <param name="args">The strongly typed object to infer factory method type parameter
    /// and to call the constructor of the infered type with arguments stored as data in this object.
    /// <para>When a value type is passed, it will be boxed because the parameter is defined
    /// as interface here. Defining the parameter with <c>in</c> would not change that.</para>
    /// </param>
    public static TWriter CreateInstance<TWriter>(IWriterArgsFor<TWriter> args)
        where TWriter : BaseWriter
    {
        var result = args.CreateInstance();
        if (result is IProvidesFinishCreateInstance finishCtor)
        {
            finishCtor.FinishCreateInstance();
        }
        return result;
    }

    //HINT: Can't use that without partial type parameter inference, because relationships in generic
    // constraints are ignored for type inference. Only types from method parameters are infered.
    // Hopefully this will also change when HKT/Type classes are added to the language.
    // see: https://github.com/dotnet/csharplang/issues/1349,
    // https://github.com/dotnet/roslyn/pull/7850,
    // https://github.com/dotnet/csharplang/issues/110
    /*public static TWriter CreateInstance<TArgs, TWriter>(ref this TArgs args)
        where TArgs : struct, IWriterArgsFor<TWriter>
        where TWriter : BaseWriter
    {
        var result = args.CreateInstance();
        if (result is IProvidesFinishCreateInstance finishCtor)
        {
            finishCtor.FinishCreateInstance();
        }
        return result;
    }*/

    public static TWriter CreateInstance<TWriter>(params object?[] args)
        where TWriter : BaseWriter
    {
        var flags = BindingFlags.CreateInstance | BindingFlags.NonPublic | BindingFlags.Public
            | BindingFlags.Instance;
        var result = (TWriter)Activator.CreateInstance(typeof(TWriter),
            flags, null, args, null)!;
        if (result is IProvidesFinishCreateInstance finishCtor)
        {
            finishCtor.FinishCreateInstance();
        }
        return result;
    }
}

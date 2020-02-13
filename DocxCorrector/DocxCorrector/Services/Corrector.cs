using System;

namespace DocxCorrector.Services
{
    public abstract class Corrector
    {
        // TODO: Remove
        public abstract void SayHi();
    }

    public sealed class CorrectorImpementation: Corrector
    {
        // TODO: Remove
        public override void SayHi() => Console.WriteLine("Hi, Corrector");
    }
}

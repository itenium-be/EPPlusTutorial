using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace EPPlusTutorial.Util
{
    /// <summary>
    /// Where can I get one?
    /// </summary>
    public class SalesGenerator
    {
        public IEnumerable<Sell> Generate(int amount)
        {
            yield return new Sell("Nails", 3.99M, 37);
            yield return new Sell(name: "Hammer", price: 12.10M, quantity: 5, discount: 0.1M);
            yield return new Sell("Saw", 15.37M, 12);
        }
    }
}

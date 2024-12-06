Function PriceAfterDiscount(originalPrice As Double, discountPercentage As Double) As Double
    ' Calculate the discounted price
    PriceAfterDiscount = originalPrice * (1 - (discountPercentage / 100))
End Function

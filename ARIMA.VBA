Option Infer On

Imports System.Data
Imports Extreme.Mathematics
Imports Extreme.Statistics
Imports Extreme.Statistics.TimeSeriesAnalysis

Namespace Extreme.Numerics.QuickStart.VB
    ' Illustrates the use of the ArimaModel class to perform
    ' estimation and forecasting of ARIMA time series models.
    Module ArimaModels

        Sub Main()
            ' This QuickStart Sample fits an ARMA(2,1) model and
            ' an ARIMA(0,1,1) model to sunspot data.

            ' The time series data is stored in a numerical variable:
            Dim sunspots = Vector.Create(
                100.8, 81.6, 66.5, 34.8, 30.6, 7, 19.8, 92.5,
                154.4, 125.9, 84.8, 68.1, 38.5, 22.8, 10.2, 24.1, 82.9,
                132, 130.9, 118.1, 89.9, 66.6, 60, 46.9, 41, 21.3, 16,
                6.4, 4.1, 6.8, 14.5, 34, 45, 43.1, 47.5, 42.2, 28.1, 10.1,
                8.1, 2.5, 0, 1.4, 5, 12.2, 13.9, 35.4, 45.8, 41.1, 30.4,
                23.9, 15.7, 6.6, 4, 1.8, 8.5, 16.6, 36.3, 49.7, 62.5, 67,
                71, 47.8, 27.5, 8.5, 13.2, 56.9, 121.5, 138.3, 103.2,
                85.8, 63.2, 36.8, 24.2, 10.7, 15, 40.1, 61.5, 98.5, 124.3,
                95.9, 66.5, 64.5, 54.2, 39, 20.6, 6.7, 4.3, 22.8, 54.8,
                93.8, 95.7, 77.2, 59.1, 44, 47, 30.5, 16.3, 7.3, 37.3,
                73.9)

            ' ARMA models (no differencing) are constructed from
            ' the variable containing the time series data, and the
            ' AR and MA orders. The following constructs an ARMA(2,1)
            ' model:
            Dim model As New ArimaModel(sunspots, 2, 1)

            ' The Compute methods fits the model.
            model.Compute()

            ' The model's Parameters collection contains the fitted values.
            ' For an ARIMA(p,d,q) model, the first p parameters are the 
            ' auto-regressive parameters. The last q parametere are the
            ' moving average parameters.
            Console.WriteLine("Variable              Value    Std.Error  t-stat  p-Value")
            For Each param As Parameter In model.Parameters
                ' Parameter objects have the following properties:
                ' - Name, usually the name of the variable:
                ' - Estimated value of the param:
                ' - Standard error:
                ' - The value of the t statistic for the hypothesis that the param
                '   is zero.
                ' - Probability corresponding to the t statistic.
                Console.WriteLine("{0,-20}{1,10:F5}{2,10:F5}{3,8:F2} {4,7:F4}", _
                    param.Name, _
                    param.Value, _
                    param.StandardError, _
                    param.Statistic, _
                    param.PValue)
            Next

            ' The log-likelihood of the computed solution is also available:
            Console.WriteLine("Log-likelihood: {0:F4}", model.LogLikelihood)
            ' as is the Akaike Information Criterion (AIC):
            Console.WriteLine("AIC: {0:F4}", model.GetAkaikeInformationCriterion())
            ' and the Baysian Information Criterion (BIC):
            Console.WriteLine("AIC: {0:F4}", model.GetBayesianInformationCriterion())

            ' The Forecast method can be used to predict the next value in the series...
            Dim nextValue As Double = model.Forecast()
            Console.WriteLine("One step ahead forecast: {0:F3}", nextValue)

            ' or to predict a specified number of values:
            Dim nextValues = model.Forecast(5)
            Console.WriteLine("First five forecasts: {0:F3}", nextValues)


            ' An integrated model (with differencing) is constructed
            ' by supplying the degree of differencing. Note the order
            ' of the orders is the traditional one for an ARIMA(p,d,q)
            ' model (p, d, q).
            ' The following constructs an ARIMA(0,1,1) model:
            Dim model2 As New ArimaModel(sunspots, 0, 1, 1)

            ' By default, the mean is assumed to be zero for an integrated model.
            ' We can override this by setting the EstimateMean property to true:
            model2.EstimateMean = True

            ' The Compute methods fits the model.
            model2.Compute()

            ' The mean shows up as one of the parameters.
            Console.WriteLine("Variable              Value    Std.Error  t-stat  p-Value")
            For Each param As Parameter In model2.Parameters
                Console.WriteLine("{0,-20}{1,10:F5}{2,10:F5}{3,8:F2} {4,7:F4}", _
                    param.Name, _
                    param.Value, _
                    param.StandardError, _
                    param.Statistic, _
                    param.PValue)
            Next

            ' We can also get the error variance:
            Console.WriteLine("Error variance: {0:F4}", model2.ErrorVariance)

            Console.WriteLine("Log-likelihood: {0:F4}", model2.LogLikelihood)
            Console.WriteLine("AIC: {0:F4}", model2.GetAkaikeInformationCriterion())
            Console.WriteLine("BIC: {0:F4}", model2.GetBayesianInformationCriterion())

            Console.WriteLine("Press Enter key to continue.")
            Console.ReadLine()
        End Sub

    End Module

End Namespace
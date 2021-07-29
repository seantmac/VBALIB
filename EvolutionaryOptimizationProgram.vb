Imports System

Namespace EvolutionaryOptimization
    Friend Class EvolutionaryOptimizationProgram
        Shared Sub Main(ByVal args() As String)
            Try
                Console.WriteLine(vbCrLf & "Begin Evolutionary Optimization demo" & vbCrLf)
                Console.WriteLine("Goal is to find the (x,y) that minimizes Schwefel's function")
                Console.WriteLine("f(x,y) = (-x * sin(sqrt(abs(x)))) + (-y * sin(sqrt(abs(y))))")
                Console.WriteLine("Known solution is x = y = 420.9687 when f = -837.9658")

                Dim popSize As Integer = 100
                Dim numGenes As Integer = 2
                Dim minGene As Double = -500.0
                Dim maxGene As Double = 500.0
                Dim mutateRate As Double = 1.0 / numGenes
                Dim precision As Double = 0.0001 ' controls mutation magnitude
                Dim tau As Double = 0.4 ' tournament selection factor
                Dim maxGeneration As Integer = 8000

                Console.WriteLine(vbCrLf & "Population size = " & popSize)
                Console.WriteLine("Number genes = " & numGenes)
                Console.WriteLine("minGene value = " & minGene.ToString("F1"))
                Console.WriteLine("maxGene value = " & maxGene.ToString("F1"))
                Console.WriteLine("Mutation rate = " & mutateRate.ToString("F4"))
                Console.WriteLine("Mutation precision = " & precision.ToString("F4"))
                Console.WriteLine("Selection pressure tau = " & tau.ToString("F2"))
                Console.WriteLine("Maximum generations = " & maxGeneration)

                Dim ev As New Evolver(popSize, numGenes, minGene, maxGene, mutateRate, precision, tau, maxGeneration) ' assumes existence of a Problem.Fitness method
                Dim best() As Double = ev.Evolve

                Console.WriteLine(vbCrLf & "Best (x,y) solution found:")
                For i = 0 To best.Length - 1
                    Console.Write(best(i).ToString("F4") & " ")
                Next i

                Dim fitness As Double = Problem.Fitness(best)
                Console.WriteLine(vbCrLf & "Function value at best solution = " & fitness.ToString("F4"))

                Console.WriteLine(vbCrLf & "End Evolutionary Optimization demo" & vbCrLf)
                Console.ReadLine()
            Catch ex As Exception
                Console.WriteLine("Fatal: " & ex.Message)
                Console.ReadLine()
            End Try
        End Sub
    End Class

    Public Class Evolver
        Private popSize As Integer
        Private population() As Individual

        Private numGenes As Integer
        Private minGene As Double
        Private maxGene As Double
        Private mutateRate As Double ' used by Mutate
        Private precision As Double ' used by Mutate

        Private tau As Double ' used by Select
        Private indexes() As Integer ' used by Select

        Private maxGeneration As Integer
        Private Shared rnd As Random = Nothing

        Public Sub New(ByVal popSize As Integer, ByVal numGenes As Integer, ByVal minGene As Double, ByVal maxGene As Double, ByVal mutateRate As Double, ByVal precision As Double, ByVal tau As Double, ByVal maxGeneration As Integer)
            Me.popSize = popSize
            Me.population = New Individual(popSize - 1) {}
            For i = 0 To population.Length - 1
                population(i) = New Individual(numGenes, minGene, maxGene, mutateRate, precision)
            Next i

            Me.numGenes = numGenes
            Me.minGene = minGene
            Me.maxGene = maxGene
            Me.mutateRate = mutateRate
            Me.precision = precision
            Me.tau = tau

            Me.indexes = New Integer(popSize - 1) {}
            For i = 0 To indexes.Length - 1
                Me.indexes(i) = i
            Next i
            Me.maxGeneration = maxGeneration
            rnd = New Random(0)
        End Sub

        Public Function Evolve() As Double()
            Dim bestFitness As Double = Me.population(0).fitness
            Dim bestChomosome(numGenes - 1) As Double
            population(0).chromosome.CopyTo(bestChomosome, 0)
            Dim gen As Integer = 0
            Do While gen < maxGeneration
                Dim parents() As Individual = [Select](2)
                Dim children() As Individual = Reproduce(parents(0), parents(1)) ' crossover & mutation
                Accept(children(0), children(1))
                Immigrate()

                For i = popSize - 3 To popSize - 1
                    If population(i).fitness < bestFitness Then
                        bestFitness = population(i).fitness
                        population(i).chromosome.CopyTo(bestChomosome, 0)
                    End If
                Next i
                gen += 1
            Loop
            Return bestChomosome
        End Function

        Private Function [Select](ByVal n As Integer) As Individual() ' select n 'good' Individuals
            'If n > popSize Then
            '    Throw New Exception("xxxx")
            'End If

            Dim tournSize As Integer = CInt(Fix(tau * popSize))
            If tournSize < n Then
                tournSize = n
            End If
            Dim candidates(tournSize - 1) As Individual

            ShuffleIndexes()
            For i = 0 To tournSize - 1
                candidates(i) = population(indexes(i))
            Next i
            Array.Sort(candidates)

            Dim results(n - 1) As Individual
            For i = 0 To n - 1
                results(i) = candidates(i)
            Next i

            Return results
        End Function

        Private Sub ShuffleIndexes()
            For i = 0 To Me.indexes.Length - 1
                Dim r As Integer = rnd.Next(i, indexes.Length)
                Dim tmp As Integer = indexes(r)
                indexes(r) = indexes(i)
                indexes(i) = tmp
            Next i
        End Sub

        'Public Overrides Function ToString() As String
        '    Dim s = New System.Text.StringBuilder
        '    For i = 0 To Me.population.Length - 1
        '        s = s.Append(i).Append(": ").Append(Me.population(i).ToString).Append(Environment.NewLine)
        '    Next (i)
        '    Return s.ToString
        'End Function

        Private Function Reproduce(ByVal parent1 As Individual, ByVal parent2 As Individual) As Individual() ' crossover and mutation
            Dim cross As Integer = rnd.Next(0, numGenes - 1) ' crossover point. 0 means 'between 0 and 1'.

            Dim child1 As New Individual(numGenes, minGene, maxGene, mutateRate, precision) ' random chromosome
            Dim child2 As New Individual(numGenes, minGene, maxGene, mutateRate, precision)

            For i = 0 To cross
                child1.chromosome(i) = parent1.chromosome(i)
            Next i
            For i = cross + 1 To numGenes - 1
                child2.chromosome(i) = parent1.chromosome(i)
            Next i
            For i = 0 To cross
                child2.chromosome(i) = parent2.chromosome(i)
            Next i
            For i = cross + 1 To numGenes - 1
                child1.chromosome(i) = parent2.chromosome(i)
            Next i

            child1.Mutate()
            child2.Mutate()

            child1.fitness = Problem.Fitness(child1.chromosome)
            child2.fitness = Problem.Fitness(child2.chromosome)

            Dim result(1) As Individual
            result(0) = child1
            result(1) = child2

            Return result
        End Function ' Reproduce

        Private Sub Accept(ByVal child1 As Individual, ByVal child2 As Individual)
            ' place child1 and chil2 into the population, replacing two worst individuals
            Array.Sort(Me.population)
            population(popSize - 1) = child1
            population(popSize - 2) = child2
            Exit Sub
        End Sub

        Private Sub Immigrate()
            Dim immigrant As New Individual(numGenes, minGene, maxGene, mutateRate, precision)
            population(popSize - 3) = immigrant ' replace third worst individual
        End Sub
    End Class ' class Evolver

    ' ------------------------------------------------------------------------------------------------

    Public Class Individual
        Implements IComparable(Of Individual)

        Public chromosome() As Double
        Public fitness As Double ' smaller values are better for minimization

        Private numGenes As Integer
        Private minGene As Double
        Private maxGene As Double
        Private mutateRate As Double
        Private precision As Double

        Private Shared rnd As New Random(0)

        Public Sub New(ByVal numGenes As Integer, ByVal minGene As Double, ByVal maxGene As Double, ByVal mutateRate As Double, ByVal precision As Double)
            Me.numGenes = numGenes
            Me.minGene = minGene
            Me.maxGene = maxGene
            Me.mutateRate = mutateRate
            Me.precision = precision
            Me.chromosome = New Double(numGenes - 1) {}
            For i = 0 To Me.chromosome.Length - 1
                Me.chromosome(i) = (maxGene - minGene) * rnd.NextDouble + minGene
            Next i
            Me.fitness = Problem.Fitness(chromosome)
        End Sub

        Public Sub Mutate()
            Dim hi As Double = precision * maxGene
            Dim lo As Double = -hi
            For i = 0 To chromosome.Length - 1
                If rnd.NextDouble < mutateRate Then
                    chromosome(i) += (hi - lo) * rnd.NextDouble + lo
                End If
            Next i
        End Sub

        'Public Overrides Function ToString() As String
        '    Dim s = New System.Text.StringBuilder
        '    For i = 0 To chromosome.Length - 1
        '        s = s.Append(chromosome(i).ToString("F2")).Append(" ")
        '    Next i
        '    If Me.fitness = Double.MaxValue Then
        '        s = s.Append("| fitness = maxValue")
        '    Else
        '        s = s.Append("| fitness = ").Append(Me.fitness.ToString("F4"))
        '    End If
        '    Return s.ToString
        'End Function

        Public Function CompareTo(ByVal other As Individual) As Integer Implements IComparable(Of Individual).CompareTo ' from smallest fitness (better) to largest
            If Me.fitness < other.fitness Then
                Return -1
            ElseIf Me.fitness > other.fitness Then
                Return 1
            Else
                Return 0
            End If
        End Function
    End Class ' class Individual

    Public Class Problem
        Public Shared Function Fitness(ByVal chromosome() As Double) As Double ' the 'cost' function we are trying to minimize
            ' Schwefel's function.
            ' for n=2, solution is x = y = 420.9687 when f(x,y) = -837.9658
            Dim result As Double = 0.0
            For i = 0 To chromosome.Length - 1
                result += (-1.0 * chromosome(i)) * Math.Sin(Math.Sqrt(Math.Abs(chromosome(i))))
            Next i
            Return result
        End Function
    End Class

    'Public Class Helpers
    '    Public Shared Sub ShowVector(ByVal vector() As Double)
    '        For i = 0 To vector.Length - 1
    '            Console.Write(vector(i).ToString("F4") & " ")
    '        Next i
    '        Console.WriteLine("")
    '    End Sub

    '    Public Shared Sub ShowVector(ByVal vector() As Integer)
    '        For i = 0 To vector.Length - 1
    '            Console.Write(vector(i) & " ")
    '        Next i
    '        Console.WriteLine("")
    '    End Sub
    'End Class
End Namespace ' ns
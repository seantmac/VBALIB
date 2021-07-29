' Demo of the k-means clustering algorithm
' No error-checking provided
' James McCaffrey, MSDN Magazine
'
' To run you can either:
' a. Launch Visual Studio and create a new VB console application,
' then zap away all template-code in Program.vb, copy the contents of this file in, and hit <F5>.
' or
' b. Copy this file to any convenient directory, launch the special Visual Studio Tools
' command shell, cd to your directory, at the prompt type csc.exe ClusteringKMeansProgram.vb and 
' hiot <enter>. You'll get an .exe in the current directory that you can run.

Namespace ClusteringKMeans
    Friend Class ClusteringKMeansProgram
        Shared Sub Main(ByVal args() As String)
            Try
                Console.WriteLine(vbCrLf & "Begin outlier data detection using k-means clustering demo" & vbCrLf)

                Console.WriteLine("Loading all (height-weight) data into memory")
                Dim attributes() As String = {"Height", "Weight"}
                Dim rawData(19)() As Double ' in most cases data will be in a text file or SQl table

                rawData(0) = New Double() {65.0, 220.0} ' if data won't fit into memory, stream through external storage
                rawData(1) = New Double() {73.0, 160.0}
                rawData(2) = New Double() {59.0, 110.0}
                rawData(3) = New Double() {61.0, 120.0}
                rawData(4) = New Double() {75.0, 150.0}
                rawData(5) = New Double() {67.0, 240.0}
                rawData(6) = New Double() {68.0, 230.0}
                rawData(7) = New Double() {70.0, 220.0}
                rawData(8) = New Double() {62.0, 130.0}
                rawData(9) = New Double() {66.0, 210.0}
                rawData(10) = New Double() {77.0, 190.0}
                rawData(11) = New Double() {75.0, 180.0}
                rawData(12) = New Double() {74.0, 170.0}
                rawData(13) = New Double() {70.0, 210.0}
                rawData(14) = New Double() {61.0, 110.0}
                rawData(15) = New Double() {58.0, 100.0}
                rawData(16) = New Double() {66.0, 230.0}
                rawData(17) = New Double() {59.0, 120.0}
                rawData(18) = New Double() {68.0, 210.0}
                rawData(19) = New Double() {61.0, 130.0}

                Console.WriteLine(vbCrLf & "Raw data:" & vbCrLf)
                ShowMatrix(rawData, rawData.Length, True)

                Dim numAttributes As Integer = attributes.Length ' 2 in this demo (height,weight)
                Dim numClusters As Integer = 3 ' vary this to experiment (must be between 2 and number data tuples)
                Dim maxCount As Integer = 30 ' trial and error

                Console.WriteLine(vbCrLf & "Begin clustering data with k = " & numClusters & " and maxCount = " & maxCount)
                Dim clustering() As Integer = Cluster(rawData, numClusters, numAttributes, maxCount)
                Console.WriteLine(vbCrLf & "Clustering complete")

                Console.WriteLine(vbCrLf & "Clustering in internal format: " & vbCrLf)
                ShowVector(clustering, True) ' true -> newline after display

                Console.WriteLine(vbCrLf & "Clustered data:")
                ShowClustering(rawData, numClusters, clustering, True)

                Dim outlier() As Double = ClusteringKMeansProgram.Outlier(rawData, clustering, numClusters, 0)
                Console.WriteLine("Outlier for cluster 0 is:")
                ShowVector(outlier, True)

                Console.WriteLine(vbCrLf & "End demo" & vbCrLf)
                Console.ReadLine()
            Catch ex As Exception
                Console.WriteLine(ex.Message)
                Console.ReadLine()
            End Try
        End Sub ' Main


        Private Shared Function InitClustering(ByVal numTuples As Integer, ByVal numClusters As Integer, ByVal randomSeed As Integer) As Integer()
            ' assign each tuple to a random cluster, making sure that there's at least
            ' one tuple assigned to every cluster
            Dim random As New Random(randomSeed)
            Dim clustering(numTuples - 1) As Integer

            ' assign first numClusters tuples to clusters 0..k-1
            For i As Integer = 0 To numClusters - 1
                clustering(i) = i
            Next i
            ' assign rest randomly
            For i As Integer = numClusters To clustering.Length - 1
                clustering(i) = random.Next(0, numClusters)
            Next i
            Return clustering
        End Function

        Private Shared Function Allocate(ByVal numClusters As Integer, ByVal numAttributes As Integer) As Double()()
            ' helper allocater for means[][] and centroids[][]
            Dim result(numClusters - 1)() As Double
            For k As Integer = 0 To numClusters - 1
                result(k) = New Double(numAttributes - 1) {}
            Next k
            Return result
        End Function

        Private Shared Sub UpdateMeans(ByVal rawData()() As Double, ByVal clustering() As Integer, ByVal means()() As Double)
            ' assumes means[][] exists. consider making means[][] a ref parameter
            Dim numClusters As Integer = means.Length
            ' zero-out means[][]
            For k As Integer = 0 To means.Length - 1
                For j As Integer = 0 To means(k).Length - 1
                    means(k)(j) = 0.0
                Next j
            Next k

            ' make an array to hold cluster counts
            Dim clusterCounts(numClusters - 1) As Integer

            ' walk through each tuple, accumulate sum for each attribute, update cluster count
            For i As Integer = 0 To rawData.Length - 1
                Dim cluster As Integer = clustering(i)
                clusterCounts(cluster) = clusterCounts(cluster) + 1
                For j As Integer = 0 To rawData(i).Length - 1
                    means(cluster)(j) = means(cluster)(j) + rawData(i)(j)
                Next j
            Next i

            ' divide each attribute sum by cluster count to get average (mean)
            For k As Integer = 0 To means.Length - 1
                For j As Integer = 0 To means(k).Length - 1
                    means(k)(j) /= clusterCounts(k) ' will throw if count is 0. consider an error-check
                Next j
            Next k

            Return
        End Sub ' UpdateMeans

        Private Shared Function ComputeCentroid(ByVal rawData()() As Double, ByVal clustering() As Integer, ByVal cluster As Integer, ByVal means()() As Double) As Double()
            ' the centroid is the actual tuple values that are closest to the cluster mean
            Dim numAttributes As Integer = means(0).Length
            Dim centroid(numAttributes - 1) As Double
            Dim minDist As Double = Double.MaxValue
            For i As Integer = 0 To rawData.Length - 1 ' walk thru each data tuple
                Dim c As Integer = clustering(i) ' if curr tuple isn't in the cluster we're computing for, continue on
                If c <> cluster Then
                    Continue For
                End If

                Dim currDist As Double = Distance(rawData(i), means(cluster)) ' call helper
                If currDist < minDist Then
                    minDist = currDist
                    For j As Integer = 0 To centroid.Length - 1
                        centroid(j) = rawData(i)(j)
                    Next j
                End If
            Next i
            Return centroid
        End Function

        Private Shared Sub UpdateCentroids(ByVal rawData()() As Double, ByVal clustering() As Integer, ByVal means()() As Double, ByVal centroids()() As Double)
            ' updates all centroids by calling helper that updates one centroid
            For k As Integer = 0 To centroids.Length - 1
                Dim centroid() As Double = ComputeCentroid(rawData, clustering, k, means)
                centroids(k) = centroid
            Next k
        End Sub

        Private Shared Function Distance(ByVal tuple() As Double, ByVal vector() As Double) As Double
            ' Euclidean distance between an actual data tuple and a cluster mean or centroid
            Dim sumSquaredDiffs As Double = 0.0
            For j As Integer = 0 To tuple.Length - 1
                sumSquaredDiffs = sumSquaredDiffs + Math.Pow((tuple(j) - vector(j)), 2)
            Next j
            Return Math.Sqrt(sumSquaredDiffs)
        End Function

        Private Shared Function MinIndex(ByVal distances() As Double) As Integer
            ' index of smallest value in distances[]
            Dim indexOfMin As Integer = 0
            Dim smallDist As Double = distances(0)
            For k As Integer = 0 To distances.Length - 1
                If distances(k) < smallDist Then
                    smallDist = distances(k)
                    indexOfMin = k
                End If
            Next k
            Return indexOfMin
        End Function

        Private Shared Function Assign(ByVal rawData()() As Double, ByVal clustering() As Integer, ByVal centroids()() As Double) As Boolean
            ' assign each tuple to best cluster (closest to cluster centroid)
            ' return true if any new cluster assignment is different from old/curr cluster
            ' does not prevent a state where a cluster has no tuples assigned. see article for details
            Dim numClusters As Integer = centroids.Length
            Dim changed As Boolean = False

            Dim distances(numClusters - 1) As Double ' distance from curr tuple to each cluster mean
            For i As Integer = 0 To rawData.Length - 1 ' walk thru each tuple
                For k As Integer = 0 To numClusters - 1 ' compute distances to all centroids
                    distances(k) = Distance(rawData(i), centroids(k))
                Next k

                Dim newCluster As Integer = MinIndex(distances) ' find the index == custerID of closest
                If newCluster <> clustering(i) Then ' different cluster assignment?
                    changed = True
                    clustering(i) = newCluster
                End If ' else no change
            Next i
            Return changed ' was there any change in clustering?
        End Function ' Assign

        Private Shared Function Cluster(ByVal rawData()() As Double, ByVal numClusters As Integer, ByVal numAttributes As Integer, ByVal maxCount As Integer) As Integer()
            Dim changed As Boolean = True
            Dim ct As Integer = 0

            Dim numTuples As Integer = rawData.Length
            Dim clustering() As Integer = InitClustering(numTuples, numClusters, 0) ' 0 is a seed for random
            Dim means()() As Double = Allocate(numClusters, numAttributes) ' just makes things a bit cleaner
            Dim centroids()() As Double = Allocate(numClusters, numAttributes)
            UpdateMeans(rawData, clustering, means) ' could call this inside UpdateCentroids instead
            UpdateCentroids(rawData, clustering, means, centroids)

            Do While changed = True AndAlso ct < maxCount
                ct = ct + 1
                changed = Assign(rawData, clustering, centroids) ' use centroids to update cluster assignment
                UpdateMeans(rawData, clustering, means) ' use new clustering to update cluster means
                UpdateCentroids(rawData, clustering, means, centroids) ' use new means to update centroids
            Loop
            'ShowMatrix(centroids, centroids.Length, true);  // show the final centroids for each cluster
            Return clustering
        End Function

        Private Shared Function Outlier(ByVal rawData()() As Double, ByVal clustering() As Integer, ByVal numClusters As Integer, ByVal cluster As Integer) As Double()
            ' return the tuple values in cluster that is farthest from cluster centroid
            Dim numAttributes As Integer = rawData(0).Length

            Dim outlier_Renamed(numAttributes - 1) As Double
            Dim maxDist As Double = 0.0

            Dim means()() As Double = Allocate(numClusters, numAttributes)
            Dim centroids()() As Double = Allocate(numClusters, numAttributes)
            UpdateMeans(rawData, clustering, means)
            UpdateCentroids(rawData, clustering, means, centroids)

            For i As Integer = 0 To rawData.Length - 1
                Dim c As Integer = clustering(i)
                If c <> cluster Then
                    Continue For
                End If
                Dim dist As Double = Distance(rawData(i), centroids(cluster))
                If dist > maxDist Then
                    maxDist = dist ' might also want to return (as an out param) the index of rawData
                    Array.Copy(rawData(i), outlier_Renamed, rawData(i).Length)
                End If
            Next i
            Return outlier_Renamed
        End Function

        ' display routines below

        Private Shared Sub ShowMatrix(ByVal matrix()() As Double, ByVal numRows As Integer, ByVal newLine As Boolean)
            For i As Integer = 0 To numRows - 1
                Console.Write("[" & i.ToString().PadLeft(2) & "]  ")
                For j As Integer = 0 To matrix(i).Length - 1
                    Console.Write(matrix(i)(j).ToString("F1") & "  ")
                Next j
                Console.WriteLine("")
            Next i
            If newLine = True Then
                Console.WriteLine("")
            End If
        End Sub ' ShowMatrix

        Private Shared Sub ShowVector(ByVal vector() As Integer, ByVal newLine As Boolean)
            For i As Integer = 0 To vector.Length - 1
                Console.Write(vector(i) & " ")
            Next i
            Console.WriteLine("")
            If newLine = True Then
                Console.WriteLine("")
            End If
        End Sub

        Private Shared Sub ShowVector(ByVal vector() As Double, ByVal newLine As Boolean)
            For i As Integer = 0 To vector.Length - 1
                Console.Write(vector(i).ToString("F1") & " ")
            Next i
            Console.WriteLine("")
            If newLine = True Then
                Console.WriteLine("")
            End If
        End Sub

        Private Shared Sub ShowClustering(ByVal rawData()() As Double, ByVal numClusters As Integer, ByVal clustering() As Integer, ByVal newLine As Boolean)
            Console.WriteLine("-----------------")
            For k As Integer = 0 To numClusters - 1 ' display by cluster
                For i As Integer = 0 To rawData.Length - 1 ' each tuple
                    If clustering(i) = k Then ' curr tuple i belongs to curr cluster k.
                        Console.Write("[" & i.ToString().PadLeft(2) & "]")
                        For j As Integer = 0 To rawData(i).Length - 1
                            Console.Write(rawData(i)(j).ToString("F1").PadLeft(6) & " ")
                        Next j
                        Console.WriteLine("")
                    End If
                Next i
                Console.WriteLine("-----------------")
            Next k
            If newLine = True Then
                Console.WriteLine("")
            End If
        End Sub

    End Class ' class
End Namespace ' ns

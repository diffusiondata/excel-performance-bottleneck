# Results

The test measured performance of updating 1,500 random rows in a 15,000-row table over 1,000 iterations, comparing two versions of Excel:

October Version (16.0.19127.20314) - degraded performance
* Total Time: 1h 5m 56.7s (3,956,707ms)
* Average Rate: 15.164 iterations/minute
* Pattern: Performance degraded sharply after the first minute (~61 iterations), then stabilized at ~14 iterations/minute

September Version (16.0.19029.20244) - baseline performance
* Total Time: 5m 17.5s (317,523ms)
* Average Rate: 188.981 iterations/minute
* Pattern: Consistent high performance throughout

*Key Finding*
The October version is ~12.5x slower than the September version for this workload. This represents a performance regression introduced in the October Excel update, specifically affecting operations that update random rows in tables. The dramatic slowdown suggests a possible issue with how the October version handles table updates, data syncing, or rendering - particularly evident in how performance dropped after the first minute and never recovered.


## Test Host

Processor: 12th Gen Intel(R) Core(TM) i7-1260P (2.10 GHz)
Installed RAM: 16.0 GB (15.7 GB usable)
Edition: Windows 11 Business
OS Build: 26100.6899

## Version 16.0.19127.20314 (October)

```
Common.ts:6 2025-10-27T10:07:33.034Z Test begun, {iterations: 1000, percentage: 10, ENV: 'Excel v16.0.19127.20314'}
Common.ts:6 2025-10-27T10:07:33.066Z Updating 1,500 random rows of table holding 15000 rows, 1,000 iterations
Common.ts:6 2025-10-27T10:08:33.075Z Iteration: 62/1000, rate: 60.993 iterations/minute
Common.ts:6 2025-10-27T10:09:33.087Z Iteration: 86/1000, rate: 23.995 iterations/minute
Common.ts:6 2025-10-27T10:10:33.101Z Iteration: 102/1000, rate: 15.997 iterations/minute
Common.ts:6 2025-10-27T10:11:33.102Z Iteration: 116/1000, rate: 14 iterations/minute
Common.ts:6 2025-10-27T10:12:33.108Z Iteration: 130/1000, rate: 13.998 iterations/minute
Common.ts:6 2025-10-27T10:13:33.113Z Iteration: 144/1000, rate: 13.999 iterations/minute
Common.ts:6 2025-10-27T10:14:33.118Z Iteration: 158/1000, rate: 13.999 iterations/minute
Common.ts:6 2025-10-27T10:15:33.132Z Iteration: 172/1000, rate: 13.997 iterations/minute
Common.ts:6 2025-10-27T10:16:33.145Z Iteration: 187/1000, rate: 14.997 iterations/minute
Common.ts:6 2025-10-27T10:17:33.159Z Iteration: 202/1000, rate: 14.997 iterations/minute
Common.ts:6 2025-10-27T10:18:33.173Z Iteration: 217/1000, rate: 14.997 iterations/minute
Common.ts:6 2025-10-27T10:19:33.185Z Iteration: 231/1000, rate: 13.997 iterations/minute
Common.ts:6 2025-10-27T10:20:33.196Z Iteration: 246/1000, rate: 14.997 iterations/minute
Common.ts:6 2025-10-27T10:21:33.198Z Iteration: 260/1000, rate: 14 iterations/minute
Common.ts:6 2025-10-27T10:22:33.212Z Iteration: 275/1000, rate: 14.996 iterations/minute
Common.ts:6 2025-10-27T10:23:33.222Z Iteration: 289/1000, rate: 13.998 iterations/minute
Common.ts:6 2025-10-27T10:24:33.238Z Iteration: 304/1000, rate: 14.996 iterations/minute
Common.ts:6 2025-10-27T10:25:33.252Z Iteration: 318/1000, rate: 13.997 iterations/minute
Common.ts:6 2025-10-27T10:26:33.265Z Iteration: 333/1000, rate: 14.997 iterations/minute
Common.ts:6 2025-10-27T10:27:33.279Z Iteration: 347/1000, rate: 13.997 iterations/minute
Common.ts:6 2025-10-27T10:28:33.293Z Iteration: 362/1000, rate: 14.997 iterations/minute
Common.ts:6 2025-10-27T10:29:33.307Z Iteration: 376/1000, rate: 13.997 iterations/minute
Common.ts:6 2025-10-27T10:30:33.320Z Iteration: 391/1000, rate: 14.997 iterations/minute
Common.ts:6 2025-10-27T10:31:33.334Z Iteration: 405/1000, rate: 13.997 iterations/minute
Common.ts:6 2025-10-27T10:32:33.348Z Iteration: 420/1000, rate: 14.997 iterations/minute
Common.ts:6 2025-10-27T10:33:33.348Z Iteration: 434/1000, rate: 14 iterations/minute
Common.ts:6 2025-10-27T10:34:33.359Z Iteration: 449/1000, rate: 14.997 iterations/minute
Common.ts:6 2025-10-27T10:35:33.371Z Iteration: 463/1000, rate: 13.997 iterations/minute
Common.ts:6 2025-10-27T10:36:33.385Z Iteration: 478/1000, rate: 14.997 iterations/minute
Common.ts:6 2025-10-27T10:37:33.399Z Iteration: 492/1000, rate: 13.997 iterations/minute
Common.ts:6 2025-10-27T10:38:33.413Z Iteration: 507/1000, rate: 14.996 iterations/minute
Common.ts:6 2025-10-27T10:39:33.426Z Iteration: 521/1000, rate: 13.997 iterations/minute
Common.ts:6 2025-10-27T10:40:33.439Z Iteration: 536/1000, rate: 14.997 iterations/minute
Common.ts:6 2025-10-27T10:41:33.453Z Iteration: 550/1000, rate: 13.997 iterations/minute
Common.ts:6 2025-10-27T10:42:33.467Z Iteration: 564/1000, rate: 13.997 iterations/minute
Common.ts:6 2025-10-27T10:43:33.482Z Iteration: 578/1000, rate: 13.997 iterations/minute
Common.ts:6 2025-10-27T10:44:33.488Z Iteration: 593/1000, rate: 14.998 iterations/minute
Common.ts:6 2025-10-27T10:45:33.492Z Iteration: 607/1000, rate: 13.999 iterations/minute
Common.ts:6 2025-10-27T10:46:33.506Z Iteration: 621/1000, rate: 13.997 iterations/minute
Common.ts:6 2025-10-27T10:47:33.508Z Iteration: 635/1000, rate: 14 iterations/minute
Common.ts:6 2025-10-27T10:48:33.516Z Iteration: 650/1000, rate: 14.998 iterations/minute
Common.ts:6 2025-10-27T10:49:33.530Z Iteration: 664/1000, rate: 13.997 iterations/minute
Common.ts:6 2025-10-27T10:50:33.543Z Iteration: 678/1000, rate: 13.997 iterations/minute
Common.ts:6 2025-10-27T10:51:33.557Z Iteration: 692/1000, rate: 13.997 iterations/minute
Common.ts:6 2025-10-27T10:52:33.570Z Iteration: 707/1000, rate: 14.997 iterations/minute
Common.ts:6 2025-10-27T10:53:33.584Z Iteration: 721/1000, rate: 13.997 iterations/minute
Common.ts:6 2025-10-27T10:54:33.598Z Iteration: 735/1000, rate: 13.997 iterations/minute
Common.ts:6 2025-10-27T10:55:33.611Z Iteration: 749/1000, rate: 13.997 iterations/minute
Common.ts:6 2025-10-27T10:56:33.626Z Iteration: 763/1000, rate: 13.997 iterations/minute
Common.ts:6 2025-10-27T10:57:33.639Z Iteration: 777/1000, rate: 13.997 iterations/minute
Common.ts:6 2025-10-27T10:58:33.653Z Iteration: 791/1000, rate: 13.997 iterations/minute
Common.ts:6 2025-10-27T10:59:33.667Z Iteration: 806/1000, rate: 14.997 iterations/minute
Common.ts:6 2025-10-27T11:00:33.680Z Iteration: 819/1000, rate: 12.997 iterations/minute
Common.ts:6 2025-10-27T11:01:33.693Z Iteration: 833/1000, rate: 13.997 iterations/minute
Common.ts:6 2025-10-27T11:02:33.707Z Iteration: 847/1000, rate: 13.997 iterations/minute
Common.ts:6 2025-10-27T11:03:33.721Z Iteration: 861/1000, rate: 13.997 iterations/minute
Common.ts:6 2025-10-27T11:04:33.727Z Iteration: 875/1000, rate: 13.999 iterations/minute
Common.ts:6 2025-10-27T11:05:33.731Z Iteration: 889/1000, rate: 13.999 iterations/minute
Common.ts:6 2025-10-27T11:06:33.739Z Iteration: 903/1000, rate: 13.998 iterations/minute
Common.ts:6 2025-10-27T11:07:33.742Z Iteration: 917/1000, rate: 13.999 iterations/minute
Common.ts:6 2025-10-27T11:08:33.756Z Iteration: 931/1000, rate: 13.997 iterations/minute
Common.ts:6 2025-10-27T11:09:33.769Z Iteration: 945/1000, rate: 13.997 iterations/minute
Common.ts:6 2025-10-27T11:10:33.778Z Iteration: 959/1000, rate: 13.998 iterations/minute
Common.ts:6 2025-10-27T11:11:33.780Z Iteration: 973/1000, rate: 14 iterations/minute
Common.ts:6 2025-10-27T11:12:33.793Z Iteration: 988/1000, rate: 14.997 iterations/minute
Common.ts:6 2025-10-27T11:13:29.776Z Completed 1,000 iterations in 3,956,707.8ms (1h 5m 56.7s), overall rate: 15.164 iterations/minute
Common.ts:6 2025-10-27T11:13:29.777Z Test completed. 1,000 iterations on 10% of rows in 3,956,742.4ms or 1h 5m 56.7s {iterations: 1000, percentage: 10}
```


## Version 16.0.19029.20244 (September)

```
Common.ts:6 2025-10-27T11:47:54.223Z Test begun, {iterations: 1000, percentage: 10, ENV: 'Excel v16.0.19029.20244'}
Common.ts:6 2025-10-27T11:47:54.252Z Updating 1,500 random rows of table holding 15000 rows, 1,000 iterations
Common.ts:6 2025-10-27T11:48:54.263Z Iteration: 182/1000, rate: 180.973 iterations/minute
Common.ts:6 2025-10-27T11:49:54.279Z Iteration: 372/1000, rate: 189.948 iterations/minute
Common.ts:6 2025-10-27T11:50:54.280Z Iteration: 559/1000, rate: 186.997 iterations/minute
Common.ts:6 2025-10-27T11:51:54.281Z Iteration: 752/1000, rate: 192.998 iterations/minute
Common.ts:6 2025-10-27T11:52:54.285Z Iteration: 945/1000, rate: 192.988 iterations/minute
Common.ts:6 2025-10-27T11:53:11.746Z Completed 1,000 iterations in 317,492ms (5m 17.4s), overall rate: 188.981 iterations/minute
Common.ts:6 2025-10-27T11:53:11.747Z Test completed. 1,000 iterations on 10% of rows in 317,523.7ms or 5m 17.5s {iterations: 1000, percentage: 10}
```


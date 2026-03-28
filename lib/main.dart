import 'dart:collection';
import 'dart:convert';
import 'dart:typed_data';

import 'package:csv/csv.dart';
import 'package:excel/excel.dart';
import 'package:file_picker/file_picker.dart';
import 'package:fl_chart/fl_chart.dart';
import 'package:flutter/material.dart';

void main() {
  runApp(const VisualizerApp());
}

const int syntheticCountColumn = -1;

class VisualizerApp extends StatefulWidget {
  const VisualizerApp({super.key});

  @override
  State<VisualizerApp> createState() => _VisualizerAppState();
}

class _VisualizerAppState extends State<VisualizerApp> {
  ThemeMode _themeMode = ThemeMode.system;
  AccentPalette _palette = AccentPalette.sunset;

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      debugShowCheckedModeBanner: false,
      title: 'Excel Visualizer',
      themeMode: _themeMode,
      theme: _buildTheme(Brightness.light, _palette.seed),
      darkTheme: _buildTheme(Brightness.dark, _palette.seed),
      home: DashboardPage(
        themeMode: _themeMode,
        palette: _palette,
        onThemeModeChanged: (mode) => setState(() => _themeMode = mode),
        onPaletteChanged: (palette) => setState(() => _palette = palette),
      ),
    );
  }
}

ThemeData _buildTheme(Brightness brightness, Color seed) {
  final scheme = ColorScheme.fromSeed(
    seedColor: seed,
    brightness: brightness,
  );
  // checking

  return ThemeData(
    useMaterial3: true,
    brightness: brightness,
    colorScheme: scheme,
    scaffoldBackgroundColor:
        brightness == Brightness.dark ? const Color(0xFF07111E) : const Color(0xFFF6F8FC),
    appBarTheme: AppBarTheme(
      centerTitle: false,
      elevation: 0,
      backgroundColor: Colors.transparent,
      foregroundColor: scheme.onSurface,
    ),
    cardTheme: CardThemeData(
      elevation: 0,
      margin: EdgeInsets.zero,
      shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(24)),
      color: brightness == Brightness.dark
          ? scheme.surface.withOpacity(0.94)
          : Colors.white.withOpacity(0.92),
    ),
    inputDecorationTheme: InputDecorationTheme(
      filled: true,
      fillColor: brightness == Brightness.dark
          ? Colors.white.withOpacity(0.06)
          : scheme.primary.withOpacity(0.06),
      border: OutlineInputBorder(
        borderRadius: BorderRadius.circular(18),
        borderSide: BorderSide.none,
      ),
      enabledBorder: OutlineInputBorder(
        borderRadius: BorderRadius.circular(18),
        borderSide: BorderSide(color: scheme.outlineVariant),
      ),
      focusedBorder: OutlineInputBorder(
        borderRadius: BorderRadius.circular(18),
        borderSide: BorderSide(color: scheme.primary, width: 1.4),
      ),
    ),
    chipTheme: ChipThemeData(
      padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 10),
      side: BorderSide(color: scheme.outlineVariant),
      shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(18)),
    ),
    segmentedButtonTheme: SegmentedButtonThemeData(
      style: ButtonStyle(
        visualDensity: VisualDensity.compact,
        side: MaterialStatePropertyAll(BorderSide(color: scheme.outlineVariant)),
        shape: MaterialStatePropertyAll(
          RoundedRectangleBorder(borderRadius: BorderRadius.circular(18)),
        ),
      ),
    ),
  );
}

class DashboardPage extends StatefulWidget {
  const DashboardPage({
    super.key,
    required this.themeMode,
    required this.palette,
    required this.onThemeModeChanged,
    required this.onPaletteChanged,
  });

  final ThemeMode themeMode;
  final AccentPalette palette;
  final ValueChanged<ThemeMode> onThemeModeChanged;
  final ValueChanged<AccentPalette> onPaletteChanged;

  @override
  State<DashboardPage> createState() => _DashboardPageState();
}

class _DashboardPageState extends State<DashboardPage> {
  ParsedDataset? _dataset;
  ChartType _chartType = ChartType.bar;
  bool _isLoading = false;
  String? _errorMessage;
  String? _fileLabel;
  int _labelColumn = 0;
  int _valueColumn = 1;
  final Set<int> _selectedSeries = {1};

  @override
  Widget build(BuildContext context) {
    final theme = Theme.of(context);
    final scheme = theme.colorScheme;
    final dataset = _dataset;

    return Scaffold(
      body: Container(
        decoration: BoxDecoration(
          gradient: LinearGradient(
            colors: [
              scheme.primary.withOpacity(theme.brightness == Brightness.dark ? 0.24 : 0.12),
              scheme.tertiary.withOpacity(theme.brightness == Brightness.dark ? 0.20 : 0.08),
              theme.scaffoldBackgroundColor,
            ],
            begin: Alignment.topLeft,
            end: Alignment.bottomRight,
          ),
        ),
        child: SafeArea(
          child: LayoutBuilder(
            builder: (context, constraints) {
              final isWide = constraints.maxWidth >= 1040;
              final sidePanel = _ControlPanel(
                dataset: dataset,
                currentChart: _chartType,
                fileLabel: _fileLabel,
                isLoading: _isLoading,
                labelColumn: _labelColumn,
                valueColumn: _valueColumn,
                selectedSeries: _selectedSeries,
                themeMode: widget.themeMode,
                palette: widget.palette,
                onUpload: _pickAndParseFile,
                onChartChanged: _onChartChanged,
                onLabelColumnChanged: _onLabelColumnChanged,
                onValueColumnChanged: _onValueColumnChanged,
                onSeriesChanged: _toggleSeriesColumn,
                onThemeModeChanged: widget.onThemeModeChanged,
                onPaletteChanged: widget.onPaletteChanged,
              );

              final mainPanel = _MainPanel(
                dataset: dataset,
                chartType: _chartType,
                errorMessage: _errorMessage,
                fileLabel: _fileLabel,
                labelColumn: _labelColumn,
                valueColumn: _valueColumn,
                selectedSeries: _selectedSeries,
              );

              return Padding(
                padding: const EdgeInsets.all(18),
                child: isWide
                    ? Row(
                        crossAxisAlignment: CrossAxisAlignment.start,
                        children: [
                          SizedBox(width: 340, child: sidePanel),
                          const SizedBox(width: 18),
                          Expanded(child: mainPanel),
                        ],
                      )
                    : SingleChildScrollView(
                        child: Column(
                          children: [
                            sidePanel,
                            const SizedBox(height: 18),
                            mainPanel,
                          ],
                        ),
                      ),
              );
            },
          ),
        ),
      ),
    );
  }

  Future<void> _pickAndParseFile() async {
    setState(() {
      _isLoading = true;
      _errorMessage = null;
    });

    try {
      final result = await FilePicker.platform.pickFiles(
        allowMultiple: false,
        withData: true,
        type: FileType.custom,
        allowedExtensions: const ['xlsx', 'csv', 'tsv', 'txt'],
      );

      if (result == null || result.files.isEmpty) {
        setState(() => _isLoading = false);
        return;
      }

      final file = result.files.single;
      final bytes = file.bytes;
      final extension = file.extension?.toLowerCase() ?? '';

      if (bytes == null || bytes.isEmpty) {
        throw const FormatException('The selected file could not be read.');
      }

      final dataset = SpreadsheetParser.parse(
        bytes: bytes,
        extension: extension,
        fileName: file.name,
      );

      final labelColumn = dataset.defaultLabelColumn;
      final valueColumn = dataset.defaultValueColumn;
      final availableValueColumns = dataset.availableValueColumns(labelColumn);
      final series = <int>{
        valueColumn,
        ...availableValueColumns.where((index) => index != valueColumn).take(3),
      };

      setState(() {
        _dataset = dataset;
        _fileLabel = file.name;
        _labelColumn = labelColumn;
        _valueColumn = valueColumn;
        _selectedSeries
          ..clear()
          ..addAll(series.take(4));
        _chartType = ChartType.bar;
        _isLoading = false;
      });
    } catch (error) {
      setState(() {
        _isLoading = false;
        _dataset = null;
        _errorMessage = error.toString().replaceFirst('FormatException: ', '');
      });
    }
  }

  void _onChartChanged(ChartType chartType) {
    setState(() => _chartType = chartType);
  }

  void _onLabelColumnChanged(int? index) {
    if (index == null) return;
    setState(() {
      _labelColumn = index;
      _selectedSeries.remove(index);
      if (_selectedSeries.isEmpty && _dataset != null) {
        _selectedSeries.addAll(_dataset!.availableValueColumns(_labelColumn).take(2));
      }
      if (_valueColumn == index) {
        _valueColumn = _nextValueColumn(_dataset, index);
      }
    });
  }

  void _onValueColumnChanged(int? index) {
    if (index == null) return;
    setState(() {
      _valueColumn = index;
      _selectedSeries.add(index);
      _selectedSeries.remove(_labelColumn);
    });
  }

  void _toggleSeriesColumn(int columnIndex) {
    if (columnIndex == _labelColumn) return;
    setState(() {
      if (_selectedSeries.contains(columnIndex)) {
        if (_selectedSeries.length > 1) {
          _selectedSeries.remove(columnIndex);
        }
      } else {
        _selectedSeries.add(columnIndex);
      }

      if (!_selectedSeries.contains(_valueColumn)) {
        _valueColumn = _selectedSeries.first;
      }
    });
  }
}

int _nextValueColumn(ParsedDataset? dataset, int labelColumn) {
  if (dataset == null) return 0;
  final available = dataset.availableValueColumns(labelColumn);
  return available.isNotEmpty ? available.first : syntheticCountColumn;
}

class _ControlPanel extends StatelessWidget {
  const _ControlPanel({
    required this.dataset,
    required this.currentChart,
    required this.fileLabel,
    required this.isLoading,
    required this.labelColumn,
    required this.valueColumn,
    required this.selectedSeries,
    required this.themeMode,
    required this.palette,
    required this.onUpload,
    required this.onChartChanged,
    required this.onLabelColumnChanged,
    required this.onValueColumnChanged,
    required this.onSeriesChanged,
    required this.onThemeModeChanged,
    required this.onPaletteChanged,
  });

  final ParsedDataset? dataset;
  final ChartType currentChart;
  final String? fileLabel;
  final bool isLoading;
  final int labelColumn;
  final int valueColumn;
  final Set<int> selectedSeries;
  final ThemeMode themeMode;
  final AccentPalette palette;
  final Future<void> Function() onUpload;
  final ValueChanged<ChartType> onChartChanged;
  final ValueChanged<int?> onLabelColumnChanged;
  final ValueChanged<int?> onValueColumnChanged;
  final ValueChanged<int> onSeriesChanged;
  final ValueChanged<ThemeMode> onThemeModeChanged;
  final ValueChanged<AccentPalette> onPaletteChanged;

  @override
  Widget build(BuildContext context) {
    final theme = Theme.of(context);
    final scheme = theme.colorScheme;

    return Column(
      children: [
        Card(
          child: Padding(
            padding: const EdgeInsets.all(22),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(
                  'Excel Visualizer',
                  style: theme.textTheme.headlineSmall?.copyWith(
                    fontWeight: FontWeight.w800,
                  ),
                ),
                const SizedBox(height: 8),
                Text(
                  'Upload `.xlsx`, `.csv`, `.tsv`, or `.txt` data and switch between pie, donut, bar, and line views instantly.',
                  style: theme.textTheme.bodyMedium?.copyWith(
                    color: scheme.onSurfaceVariant,
                    height: 1.45,
                  ),
                ),
                const SizedBox(height: 18),
                FilledButton.icon(
                  onPressed: isLoading ? null : onUpload,
                  icon: isLoading
                      ? const SizedBox(
                          width: 16,
                          height: 16,
                          child: CircularProgressIndicator(strokeWidth: 2),
                        )
                      : const Icon(Icons.upload_file_rounded),
                  label: Text(isLoading ? 'Reading file...' : 'Upload spreadsheet'),
                ),
                if (fileLabel != null) ...[
                  const SizedBox(height: 14),
                  Container(
                    padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 10),
                    decoration: BoxDecoration(
                      color: scheme.primary.withOpacity(0.08),
                      borderRadius: BorderRadius.circular(16),
                    ),
                    child: Row(
                      children: [
                        Icon(Icons.description_outlined, color: scheme.primary),
                        const SizedBox(width: 10),
                        Expanded(
                          child: Text(
                            fileLabel!,
                            overflow: TextOverflow.ellipsis,
                            style: theme.textTheme.bodyMedium?.copyWith(
                              fontWeight: FontWeight.w600,
                            ),
                          ),
                        ),
                      ],
                    ),
                  ),
                ],
              ],
            ),
          ),
        ),
        const SizedBox(height: 18),
        Card(
          child: Padding(
            padding: const EdgeInsets.all(22),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(
                  'Theme buttons',
                  style: theme.textTheme.titleMedium?.copyWith(fontWeight: FontWeight.w700),
                ),
                const SizedBox(height: 14),
                SegmentedButton<ThemeMode>(
                  segments: const [
                    ButtonSegment(
                      value: ThemeMode.light,
                      icon: Icon(Icons.light_mode_rounded),
                      label: Text('Light'),
                    ),
                    ButtonSegment(
                      value: ThemeMode.system,
                      icon: Icon(Icons.brightness_auto_rounded),
                      label: Text('Auto'),
                    ),
                    ButtonSegment(
                      value: ThemeMode.dark,
                      icon: Icon(Icons.dark_mode_rounded),
                      label: Text('Dark'),
                    ),
                  ],
                  selected: {themeMode},
                  onSelectionChanged: (selection) {
                    if (selection.isNotEmpty) {
                      onThemeModeChanged(selection.first);
                    }
                  },
                ),
                const SizedBox(height: 18),
                Wrap(
                  spacing: 10,
                  runSpacing: 10,
                  children: AccentPalette.values.map((entry) {
                    final isSelected = entry == palette;
                    return FilterChip(
                      selected: isSelected,
                      label: Text(entry.label),
                      avatar: CircleAvatar(backgroundColor: entry.seed, radius: 9),
                      onSelected: (_) => onPaletteChanged(entry),
                    );
                  }).toList(),
                ),
              ],
            ),
          ),
        ),
        const SizedBox(height: 18),
        Card(
          child: Padding(
            padding: const EdgeInsets.all(22),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(
                  'Visualization',
                  style: theme.textTheme.titleMedium?.copyWith(fontWeight: FontWeight.w700),
                ),
                const SizedBox(height: 14),
                SegmentedButton<ChartType>(
                  segments: const [
                    ButtonSegment(
                      value: ChartType.bar,
                      icon: Icon(Icons.bar_chart_rounded),
                      label: Text('Bar'),
                    ),
                    ButtonSegment(
                      value: ChartType.line,
                      icon: Icon(Icons.show_chart_rounded),
                      label: Text('Line'),
                    ),
                    ButtonSegment(
                      value: ChartType.pie,
                      icon: Icon(Icons.pie_chart_rounded),
                      label: Text('Pie'),
                    ),
                    ButtonSegment(
                      value: ChartType.donut,
                      icon: Icon(Icons.donut_large_rounded),
                      label: Text('Donut'),
                    ),
                  ],
                  selected: {currentChart},
                  onSelectionChanged: (selection) {
                    if (selection.isNotEmpty) {
                      onChartChanged(selection.first);
                    }
                  },
                ),
                const SizedBox(height: 18),
                if (dataset != null) ...[
                  ...() {
                    final valueOptions = dataset!.availableValueColumns(labelColumn);
                    final safeValueColumn = valueOptions.contains(valueColumn)
                        ? valueColumn
                        : valueOptions.first;

                    return [
                  DropdownButtonFormField<int>(
                    value: labelColumn,
                    decoration: const InputDecoration(labelText: 'Label column'),
                    items: dataset!.headers
                        .asMap()
                        .entries
                        .map(
                          (entry) => DropdownMenuItem(
                            value: entry.key,
                            child: Text(entry.value),
                          ),
                        )
                        .toList(),
                    onChanged: onLabelColumnChanged,
                  ),
                  const SizedBox(height: 14),
                  DropdownButtonFormField<int>(
                    value: safeValueColumn,
                    decoration: const InputDecoration(labelText: 'Primary value column'),
                    items: valueOptions
                        .map(
                          (index) => DropdownMenuItem(
                            value: index,
                            child: Text(dataset!.displayNameForColumn(index)),
                          ),
                        )
                        .toList(),
                    onChanged: onValueColumnChanged,
                  ),
                  const SizedBox(height: 16),
                  Text(
                    'Visible series',
                    style: theme.textTheme.labelLarge?.copyWith(fontWeight: FontWeight.w700),
                  ),
                  const SizedBox(height: 10),
                  Wrap(
                    spacing: 10,
                    runSpacing: 10,
                    children: valueOptions
                        .map(
                          (index) => FilterChip(
                            selected: selectedSeries.contains(index),
                            label: Text(dataset!.displayNameForColumn(index)),
                            onSelected: (_) => onSeriesChanged(index),
                          ),
                        )
                        .toList(),
                  ),
                    ];
                  }(),
                ] else
                  Text(
                    'Once a file is loaded, you can choose which columns drive the labels and chart values.',
                    style: theme.textTheme.bodyMedium?.copyWith(
                      color: scheme.onSurfaceVariant,
                      height: 1.45,
                    ),
                  ),
              ],
            ),
          ),
        ),
      ],
    );
  }
}

class _MainPanel extends StatelessWidget {
  const _MainPanel({
    required this.dataset,
    required this.chartType,
    required this.errorMessage,
    required this.fileLabel,
    required this.labelColumn,
    required this.valueColumn,
    required this.selectedSeries,
  });

  final ParsedDataset? dataset;
  final ChartType chartType;
  final String? errorMessage;
  final String? fileLabel;
  final int labelColumn;
  final int valueColumn;
  final Set<int> selectedSeries;

  @override
  Widget build(BuildContext context) {
    final theme = Theme.of(context);
    final scheme = theme.colorScheme;

    if (errorMessage != null) {
      return _MessageCard(
        icon: Icons.warning_amber_rounded,
        title: 'File could not be visualized',
        message: errorMessage!,
      );
    }

    if (dataset == null) {
      return const _MessageCard(
        icon: Icons.insights_rounded,
        title: 'Drop in a spreadsheet to begin',
        message:
            'Your chart studio will appear here with summary cards, interactive graphs, and a preview table after upload.',
      );
    }

    final chartSeries = dataset!.buildSeries(
      labelColumn: labelColumn,
      valueColumn: valueColumn,
      selectedColumns: selectedSeries,
    );

    return Column(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: [
        Row(
          children: [
            Expanded(
              child: Text(
                'Data studio',
                style: theme.textTheme.headlineSmall?.copyWith(fontWeight: FontWeight.w800),
              ),
            ),
            if (fileLabel != null)
              Chip(
                label: Text(fileLabel!),
                avatar: const Icon(Icons.table_chart_rounded, size: 18),
              ),
          ],
        ),
        const SizedBox(height: 16),
        Wrap(
          spacing: 16,
          runSpacing: 16,
          children: [
            _StatCard(
              title: 'Rows',
              value: '${dataset!.rows.length}',
              caption: 'Records imported',
            ),
            _StatCard(
              title: 'Columns',
              value: '${dataset!.headers.length}',
              caption: 'Detected headers',
            ),
            _StatCard(
              title: 'Chart fields',
              value: '${dataset!.availableValueColumns(labelColumn).length}',
              caption: 'Usable values and count mode',
            ),
            _StatCard(
              title: 'Chart type',
              value: chartType.label,
              caption: 'Live visualization mode',
            ),
          ],
        ),
        const SizedBox(height: 18),
        Card(
          child: Padding(
            padding: const EdgeInsets.all(20),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Row(
                  children: [
                    Expanded(
                      child: Text(
                        '${chartType.label} chart',
                        style: theme.textTheme.titleLarge?.copyWith(fontWeight: FontWeight.w700),
                      ),
                    ),
                    Text(
                      'Labels: ${dataset!.headers[labelColumn]}',
                      style: theme.textTheme.bodyMedium?.copyWith(color: scheme.onSurfaceVariant),
                    ),
                  ],
                ),
                const SizedBox(height: 18),
                SizedBox(
                  height: 360,
                  child: ChartViewport(
                    chartType: chartType,
                    seriesCollection: chartSeries,
                    colorScheme: scheme,
                  ),
                ),
              ],
            ),
          ),
        ),
        const SizedBox(height: 18),
        Card(
          child: Padding(
            padding: const EdgeInsets.all(20),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(
                  'Preview table',
                  style: theme.textTheme.titleLarge?.copyWith(fontWeight: FontWeight.w700),
                ),
                const SizedBox(height: 14),
                SingleChildScrollView(
                  scrollDirection: Axis.horizontal,
                  child: DataTable(
                    headingRowHeight: 46,
                    dataRowMinHeight: 42,
                    dataRowMaxHeight: 56,
                    columns: dataset!.headers
                        .map(
                          (header) => DataColumn(
                            label: Text(
                              header,
                              style: theme.textTheme.labelLarge?.copyWith(
                                fontWeight: FontWeight.w700,
                              ),
                            ),
                          ),
                        )
                        .toList(),
                    rows: dataset!.rows
                        .take(10)
                        .map(
                          (row) => DataRow(
                            cells: row.values
                                .map((value) => DataCell(Text(value.display)))
                                .toList(),
                          ),
                        )
                        .toList(),
                  ),
                ),
              ],
            ),
          ),
        ),
      ],
    );
  }
}

class ChartViewport extends StatelessWidget {
  const ChartViewport({
    super.key,
    required this.chartType,
    required this.seriesCollection,
    required this.colorScheme,
  });

  final ChartType chartType;
  final List<ChartSeries> seriesCollection;
  final ColorScheme colorScheme;

  @override
  Widget build(BuildContext context) {
    if (seriesCollection.isEmpty || seriesCollection.first.points.isEmpty) {
      return Center(
        child: Text(
          'This file needs at least one label column and one numeric column to draw a chart.',
          textAlign: TextAlign.center,
          style: Theme.of(context).textTheme.bodyLarge,
        ),
      );
    }

    switch (chartType) {
      case ChartType.bar:
        return _buildBarChart(context);
      case ChartType.line:
        return _buildLineChart(context);
      case ChartType.pie:
        return _buildPieChart(context, isDonut: false);
      case ChartType.donut:
        return _buildPieChart(context, isDonut: true);
    }
  }

  Widget _buildBarChart(BuildContext context) {
    final legend = _Legend(seriesCollection: seriesCollection);
    final maxValue = _maxY(seriesCollection);

    return Column(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: [
        if (seriesCollection.length > 1) legend,
        if (seriesCollection.length > 1) const SizedBox(height: 12),
        Expanded(
          child: BarChart(
            BarChartData(
              alignment: BarChartAlignment.spaceAround,
              maxY: maxValue == 0 ? 10 : maxValue * 1.2,
              gridData: FlGridData(
                drawVerticalLine: false,
                horizontalInterval: maxValue <= 5 ? 1 : maxValue / 5,
                getDrawingHorizontalLine: (value) => FlLine(
                  color: colorScheme.outlineVariant.withOpacity(0.4),
                  strokeWidth: 1,
                ),
              ),
              borderData: FlBorderData(show: false),
              titlesData: FlTitlesData(
                topTitles: const AxisTitles(sideTitles: SideTitles(showTitles: false)),
                rightTitles: const AxisTitles(sideTitles: SideTitles(showTitles: false)),
                leftTitles: AxisTitles(
                  sideTitles: SideTitles(
                    showTitles: true,
                    reservedSize: 42,
                    getTitlesWidget: (value, _) => Text(
                      value.toStringAsFixed(value % 1 == 0 ? 0 : 1),
                      style: Theme.of(context).textTheme.labelSmall,
                    ),
                  ),
                ),
                bottomTitles: AxisTitles(
                  sideTitles: SideTitles(
                    showTitles: true,
                    reservedSize: 42,
                    getTitlesWidget: (value, _) {
                      final index = value.toInt();
                      if (index < 0 || index >= seriesCollection.first.points.length) {
                        return const SizedBox.shrink();
                      }
                      return Padding(
                        padding: const EdgeInsets.only(top: 10),
                        child: Text(
                          seriesCollection.first.points[index].label,
                          overflow: TextOverflow.ellipsis,
                          style: Theme.of(context).textTheme.labelSmall,
                        ),
                      );
                    },
                  ),
                ),
              ),
              barGroups: List.generate(seriesCollection.first.points.length, (pointIndex) {
                final rods = <BarChartRodData>[];
                for (var seriesIndex = 0; seriesIndex < seriesCollection.length; seriesIndex++) {
                  final series = seriesCollection[seriesIndex];
                  rods.add(
                    BarChartRodData(
                      toY: series.points[pointIndex].value,
                      width: seriesCollection.length == 1 ? 28 : 14,
                      borderRadius: BorderRadius.circular(6),
                      color: series.color,
                    ),
                  );
                }
                return BarChartGroupData(x: pointIndex, barRods: rods, barsSpace: 8);
              }),
            ),
          ),
        ),
      ],
    );
  }

  Widget _buildLineChart(BuildContext context) {
    final maxValue = _maxY(seriesCollection);

    return Column(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: [
        _Legend(seriesCollection: seriesCollection),
        const SizedBox(height: 12),
        Expanded(
          child: LineChart(
            LineChartData(
              minY: 0,
              maxY: maxValue == 0 ? 10 : maxValue * 1.2,
              gridData: FlGridData(
                drawVerticalLine: false,
                horizontalInterval: maxValue <= 5 ? 1 : maxValue / 5,
                getDrawingHorizontalLine: (value) => FlLine(
                  color: colorScheme.outlineVariant.withOpacity(0.4),
                  strokeWidth: 1,
                ),
              ),
              borderData: FlBorderData(show: false),
              titlesData: FlTitlesData(
                topTitles: const AxisTitles(sideTitles: SideTitles(showTitles: false)),
                rightTitles: const AxisTitles(sideTitles: SideTitles(showTitles: false)),
                leftTitles: AxisTitles(
                  sideTitles: SideTitles(
                    showTitles: true,
                    reservedSize: 42,
                    getTitlesWidget: (value, _) => Text(
                      value.toStringAsFixed(value % 1 == 0 ? 0 : 1),
                      style: Theme.of(context).textTheme.labelSmall,
                    ),
                  ),
                ),
                bottomTitles: AxisTitles(
                  sideTitles: SideTitles(
                    showTitles: true,
                    reservedSize: 42,
                    getTitlesWidget: (value, _) {
                      final index = value.toInt();
                      if (index < 0 || index >= seriesCollection.first.points.length) {
                        return const SizedBox.shrink();
                      }
                      return Padding(
                        padding: const EdgeInsets.only(top: 10),
                        child: Text(
                          seriesCollection.first.points[index].label,
                          overflow: TextOverflow.ellipsis,
                          style: Theme.of(context).textTheme.labelSmall,
                        ),
                      );
                    },
                  ),
                ),
              ),
              lineBarsData: seriesCollection.map((series) {
                return LineChartBarData(
                  isCurved: true,
                  color: series.color,
                  barWidth: 3,
                  belowBarData: BarAreaData(
                    show: true,
                    color: series.color.withOpacity(0.12),
                  ),
                  dotData: FlDotData(show: series.points.length <= 20),
                  spots: List.generate(
                    series.points.length,
                    (index) => FlSpot(index.toDouble(), series.points[index].value),
                  ),
                );
              }).toList(),
            ),
          ),
        ),
      ],
    );
  }

  Widget _buildPieChart(BuildContext context, {required bool isDonut}) {
    final primarySeries = seriesCollection.first;
    final total = primarySeries.points.fold<double>(0, (sum, point) => sum + point.value);

    return Row(
      children: [
        Expanded(
          flex: 6,
          child: PieChart(
            PieChartData(
              centerSpaceRadius: isDonut ? 54 : 0,
              sectionsSpace: 3,
              sections: primarySeries.points.map((point) {
                final percentage = total == 0 ? 0 : (point.value / total) * 100;
                return PieChartSectionData(
                  color: point.color,
                  radius: isDonut ? 86 : 96,
                  title: '${percentage.toStringAsFixed(1)}%',
                  value: point.value,
                  titleStyle: Theme.of(context).textTheme.labelSmall?.copyWith(
                        color: Colors.white,
                        fontWeight: FontWeight.w700,
                      ),
                );
              }).toList(),
            ),
          ),
        ),
        const SizedBox(width: 18),
        Expanded(
          flex: 5,
          child: ListView.separated(
            itemCount: primarySeries.points.length,
            separatorBuilder: (_, __) => const Divider(height: 18),
            itemBuilder: (context, index) {
              final point = primarySeries.points[index];
              final percentage = total == 0 ? 0 : (point.value / total) * 100;
              return Row(
                children: [
                  Container(
                    width: 12,
                    height: 12,
                    decoration: BoxDecoration(
                      color: point.color,
                      shape: BoxShape.circle,
                    ),
                  ),
                  const SizedBox(width: 12),
                  Expanded(
                    child: Text(
                      point.label,
                      overflow: TextOverflow.ellipsis,
                      style: Theme.of(context).textTheme.bodyMedium?.copyWith(
                            fontWeight: FontWeight.w600,
                          ),
                    ),
                  ),
                  const SizedBox(width: 10),
                  Text(
                    '${point.value.toStringAsFixed(1)} (${percentage.toStringAsFixed(1)}%)',
                    style: Theme.of(context).textTheme.bodySmall,
                  ),
                ],
              );
            },
          ),
        ),
      ],
    );
  }

  double _maxY(List<ChartSeries> seriesCollection) {
    double max = 0;
    for (final series in seriesCollection) {
      for (final point in series.points) {
        if (point.value > max) {
          max = point.value;
        }
      }
    }
    return max;
  }
}

class _Legend extends StatelessWidget {
  const _Legend({required this.seriesCollection});

  final List<ChartSeries> seriesCollection;

  @override
  Widget build(BuildContext context) {
    return Wrap(
      spacing: 12,
      runSpacing: 12,
      children: seriesCollection.map((series) {
        return Row(
          mainAxisSize: MainAxisSize.min,
          children: [
            Container(
              width: 12,
              height: 12,
              decoration: BoxDecoration(
                color: series.color,
                borderRadius: BorderRadius.circular(4),
              ),
            ),
            const SizedBox(width: 8),
            Text(series.name, style: Theme.of(context).textTheme.bodySmall),
          ],
        );
      }).toList(),
    );
  }
}

class _MessageCard extends StatelessWidget {
  const _MessageCard({
    required this.icon,
    required this.title,
    required this.message,
  });

  final IconData icon;
  final String title;
  final String message;

  @override
  Widget build(BuildContext context) {
    final theme = Theme.of(context);
    final scheme = theme.colorScheme;

    return Card(
      child: SizedBox(
        height: 420,
        child: Center(
          child: Padding(
            padding: const EdgeInsets.all(28),
            child: Column(
              mainAxisSize: MainAxisSize.min,
              children: [
                Container(
                  width: 72,
                  height: 72,
                  decoration: BoxDecoration(
                    color: scheme.primary.withOpacity(0.10),
                    borderRadius: BorderRadius.circular(24),
                  ),
                  child: Icon(icon, size: 36, color: scheme.primary),
                ),
                const SizedBox(height: 18),
                Text(
                  title,
                  textAlign: TextAlign.center,
                  style: theme.textTheme.headlineSmall?.copyWith(fontWeight: FontWeight.w800),
                ),
                const SizedBox(height: 10),
                Text(
                  message,
                  textAlign: TextAlign.center,
                  style: theme.textTheme.bodyLarge?.copyWith(
                    color: scheme.onSurfaceVariant,
                    height: 1.5,
                  ),
                ),
              ],
            ),
          ),
        ),
      ),
    );
  }
}

class _StatCard extends StatelessWidget {
  const _StatCard({
    required this.title,
    required this.value,
    required this.caption,
  });

  final String title;
  final String value;
  final String caption;

  @override
  Widget build(BuildContext context) {
    final theme = Theme.of(context);

    return Card(
      child: SizedBox(
        width: 190,
        child: Padding(
          padding: const EdgeInsets.all(18),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              Text(title, style: theme.textTheme.labelLarge),
              const SizedBox(height: 10),
              Text(
                value,
                style: theme.textTheme.headlineMedium?.copyWith(fontWeight: FontWeight.w800),
              ),
              const SizedBox(height: 6),
              Text(caption, style: theme.textTheme.bodySmall),
            ],
          ),
        ),
      ),
    );
  }
}

enum ChartType {
  bar('Bar'),
  line('Line'),
  pie('Pie'),
  donut('Donut');

  const ChartType(this.label);
  final String label;
}

enum AccentPalette {
  sunset('Sunset', Color(0xFFE96D71)),
  ocean('Ocean', Color(0xFF2F7CF6)),
  forest('Forest', Color(0xFF1E9B6C)),
  ember('Ember', Color(0xFFDA7B27));

  const AccentPalette(this.label, this.seed);
  final String label;
  final Color seed;
}

class SpreadsheetParser {
  static ParsedDataset parse({
    required Uint8List bytes,
    required String extension,
    required String fileName,
  }) {
    final normalized = extension.toLowerCase();

    if (normalized == 'xlsx') {
      return _parseExcel(bytes, fileName);
    }

    if (normalized == 'csv' || normalized == 'txt' || normalized == 'tsv') {
      return _parseDelimited(bytes, fileName, delimiter: normalized == 'tsv' ? '\t' : ',');
    }

    throw const FormatException('Please upload a supported file: .xlsx, .csv, .tsv, or .txt');
  }

  static ParsedDataset _parseDelimited(
    Uint8List bytes,
    String fileName, {
    required String delimiter,
  }) {
    final decoded = const Utf8Decoder(allowMalformed: true).convert(bytes);
    final effectiveDelimiter = _detectDelimiter(decoded, preferred: delimiter);
    final rows = const CsvToListConverter(
      shouldParseNumbers: false,
      eol: '\n',
    ).convert(decoded, fieldDelimiter: effectiveDelimiter);

    return _datasetFromRows(
      fileName: fileName,
      rows: rows.map((row) => row.map((cell) => cell?.toString() ?? '').toList()).toList(),
    );
  }

  static ParsedDataset _parseExcel(Uint8List bytes, String fileName) {
    try {
      final workbook = Excel.decodeBytes(bytes);
      if (workbook.tables.isEmpty) {
        throw const FormatException('No sheets were found in the uploaded workbook.');
      }

      final firstSheet = workbook.tables.values.first;
      final rows = <List<String>>[];

      for (final row in firstSheet.rows) {
        rows.add(row.map(_cellToString).toList());
      }

      return _datasetFromRows(fileName: fileName, rows: rows);
    } catch (_) {
      throw const FormatException(
        'This spreadsheet could not be parsed safely. Try saving it again as .xlsx or .csv and re-uploading.',
      );
    }
  }

  static String _detectDelimiter(String decoded, {required String preferred}) {
    final sampleLines = decoded
        .split(RegExp(r'\r?\n'))
        .where((line) => line.trim().isNotEmpty)
        .take(5)
        .toList();
    if (sampleLines.isEmpty) return preferred;

    const candidates = [',', ';', '\t', '|'];
    String best = preferred;
    var bestScore = -1;

    for (final candidate in candidates) {
      final score = sampleLines.fold<int>(
        0,
        (sum, line) => sum + candidate.allMatches(line).length,
      );
      if (score > bestScore) {
        bestScore = score;
        best = candidate;
      }
    }

    return bestScore > 0 ? best : preferred;
  }

  static String _cellToString(dynamic cell) {
    if (cell == null) return '';
    try {
      final dynamic value = cell.value;
      if (value == null) return '';
      return value.toString();
    } catch (_) {
      try {
        return cell.toString();
      } catch (_) {
        return '';
      }
    }
  }

  static ParsedDataset _datasetFromRows({
    required String fileName,
    required List<List<String>> rows,
  }) {
    if (rows.isEmpty) {
      throw const FormatException('The selected file is empty.');
    }

    final normalizedRows = rows
        .where((row) => row.any((cell) => cell.trim().isNotEmpty))
        .map((row) => row.map((cell) => cell.trim()).toList())
        .toList();

    if (normalizedRows.isEmpty) {
      throw const FormatException('The selected file does not contain readable rows.');
    }

    final width = normalizedRows.map((row) => row.length).reduce((a, b) => a > b ? a : b);
    final paddedRows = normalizedRows
        .map(
          (row) => [
            ...row,
            ...List.filled(width - row.length, ''),
          ],
        )
        .toList();

    final firstRowIsHeader = _looksLikeHeaderRow(paddedRows);
    final headers = (firstRowIsHeader ? paddedRows.first : List<String>.generate(width, (index) => 'Column ${index + 1}'))
        .asMap()
        .entries
        .map((entry) => entry.value.isEmpty ? 'Column ${entry.key + 1}' : entry.value)
        .toList();
    final dataRows = firstRowIsHeader ? paddedRows.skip(1).toList() : paddedRows;
    final records = dataRows.map((row) => ParsedRow.fromStrings(row)).toList();

    if (records.isEmpty) {
      throw const FormatException('The selected file does not contain any data rows.');
    }

    final numericIndexes = <int>[];
    for (var columnIndex = 0; columnIndex < headers.length; columnIndex++) {
      final values = records.map((row) => row.values[columnIndex]).where((value) => value.display.isNotEmpty);
      final numericMatches = values.where((value) => value.numeric != null).length;
      if (numericMatches > 0) {
        numericIndexes.add(columnIndex);
      }
    }

    final defaultLabelColumn = headers.asMap().keys.firstWhere(
          (index) => !numericIndexes.contains(index),
          orElse: () => 0,
        );
    final defaultValueColumn = numericIndexes.firstWhere(
          (index) => index != defaultLabelColumn,
          orElse: () => syntheticCountColumn,
        );

    return ParsedDataset(
      fileName: fileName,
      headers: headers,
      rows: records,
      numericColumnIndexes: numericIndexes,
      defaultLabelColumn: defaultLabelColumn,
      defaultValueColumn: defaultValueColumn,
    );
  }

  static bool _looksLikeHeaderRow(List<List<String>> rows) {
    if (rows.length < 2) return false;

    final first = rows.first;
    final second = rows[1];
    var firstNumeric = 0;
    var secondNumeric = 0;
    var firstText = 0;

    for (var index = 0; index < first.length; index++) {
      final firstCell = ParsedCell.fromRaw(first[index]);
      final secondCell = ParsedCell.fromRaw(second[index]);
      if (firstCell.numeric != null) firstNumeric++;
      if (secondCell.numeric != null) secondNumeric++;
      if (firstCell.display.isNotEmpty && firstCell.numeric == null) firstText++;
    }

    return firstText >= firstNumeric || secondNumeric > firstNumeric;
  }
}

class ParsedDataset {
  const ParsedDataset({
    required this.fileName,
    required this.headers,
    required this.rows,
    required this.numericColumnIndexes,
    required this.defaultLabelColumn,
    required this.defaultValueColumn,
  });

  final String fileName;
  final List<String> headers;
  final List<ParsedRow> rows;
  final List<int> numericColumnIndexes;
  final int defaultLabelColumn;
  final int defaultValueColumn;

  List<int> availableValueColumns(int labelColumn) {
    final values = <int>[
      syntheticCountColumn,
      ...numericColumnIndexes.where((index) => index != labelColumn),
    ];

    return LinkedHashSet<int>.from(values).toList();
  }

  String displayNameForColumn(int columnIndex) {
    if (columnIndex == syntheticCountColumn) {
      return 'Row count';
    }
    return headers[columnIndex];
  }

  List<ChartSeries> buildSeries({
    required int labelColumn,
    required int valueColumn,
    required Set<int> selectedColumns,
  }) {
    final palette = [
      const Color(0xFF2F7CF6),
      const Color(0xFFE96D71),
      const Color(0xFF13A37F),
      const Color(0xFFF2A93B),
      const Color(0xFF845EF7),
      const Color(0xFF0EA5E9),
    ];

    final activeColumns = selectedColumns.where((index) => index == syntheticCountColumn || index != labelColumn).toList();
    if (activeColumns.isEmpty) {
      activeColumns.add(valueColumn);
    }

    return activeColumns.asMap().entries.map((entry) {
      final seriesIndex = entry.key;
      final columnIndex = entry.value;
      final color = palette[seriesIndex % palette.length];
      final groupedPoints = LinkedHashMap<String, double>();

      for (var rowIndex = 0; rowIndex < rows.length; rowIndex++) {
        final row = rows[rowIndex];
        final label = row.values[labelColumn].display.isEmpty
            ? 'Item ${rowIndex + 1}'
            : row.values[labelColumn].display;
        final value = columnIndex == syntheticCountColumn
            ? 1.0
            : (row.values[columnIndex].numeric ?? 0);

        groupedPoints.update(label, (existing) => existing + value, ifAbsent: () => value);
      }

      final points = groupedPoints.entries
          .map(
            (point) => ChartPoint(
              label: point.key,
              value: point.value,
              color: color,
            ),
          )
          .toList();

      return ChartSeries(
        name: displayNameForColumn(columnIndex),
        color: color,
        points: points,
      );
    }).toList();
  }
}

class ParsedRow {
  const ParsedRow(this.values);

  final List<ParsedCell> values;

  factory ParsedRow.fromStrings(List<String> row) {
    return ParsedRow(row.map(ParsedCell.fromRaw).toList());
  }
}

class ParsedCell {
  const ParsedCell({
    required this.display,
    required this.numeric,
  });

  final String display;
  final double? numeric;

  static final RegExp _sanitizePattern = RegExp(r'[^0-9.\-]');

  factory ParsedCell.fromRaw(String raw) {
    final cleaned = raw.trim();
    final normalized = cleaned.replaceAll(',', '');
    final numeric = double.tryParse(normalized) ??
        double.tryParse(normalized.replaceAll(_sanitizePattern, ''));

    return ParsedCell(
      display: cleaned,
      numeric: numeric,
    );
  }
}

class ChartSeries {
  const ChartSeries({
    required this.name,
    required this.color,
    required this.points,
  });

  final String name;
  final Color color;
  final List<ChartPoint> points;
}

class ChartPoint {
  const ChartPoint({
    required this.label,
    required this.value,
    required this.color,
  });

  final String label;
  final double value;
  final Color color;
}

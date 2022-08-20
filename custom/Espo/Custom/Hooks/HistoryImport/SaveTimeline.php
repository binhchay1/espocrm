<?php

namespace Espo\Custom\Hooks\HistoryImport;

use Espo\ORM\Entity;
use Espo\ORM\EntityManager;
use Espo\Core\Di;

class SaveTimeline implements
    Di\FileManagerAware,
    Di\FileStorageManagerAware,
    Di\ConfigAware
{
    use Di\FileManagerSetter;
    use Di\FileStorageManagerSetter;
    use Di\ConfigSetter;

    protected $entityManager;
    protected $entity;

    public function __construct(EntityManager $entityManager, Entity $entity)
    {
        $this->entityManager = $entityManager;
        $this->entity = $entity;
    }

    public function afterCreateEntity(Entity $entity, array $options, array $data)
    {
        $id = $entity->get('id');
        $defaultStorage = $this->config->get('defaultFileStorage');

        if ($defaultStorage == 'EspoUploadDir') {
            $storage = 'data/upload/';
        } else {
            $storage = $defaultStorage;
        }

        $request_body = file_get_contents('php://input');
        $data = json_decode($request_body);
        $fileName = $data->fileName;
        $pathFileDefault = $storage . $id;
        $pathFileNew = $storage . $fileName;
        copy($pathFileDefault, $pathFileNew);
        $inputFileName = $pathFileNew;

        // $inputFileName = $storage . 'LLV NH BTQM 3 T8.xlsx';
        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $sheet = $reader->load($inputFileName)->getActiveSheet();
        $this->readFile($sheet);

        unlink($pathFileNew);
    }

    public function readFile($worksheet)
    {
        $highestRow = $worksheet->getHighestRow();
        $highestColumn = $worksheet->getHighestColumn();
        $highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn);
        $timeLineType = $entityManager
            ->getRDBRepository('TimeLineType')
            ->find();

        $ignoreHeader = ['thu ngân', 'bql', 'bar', 'lễ tân', 'nhân viên phục vụ', 'đại sứ'];
        $arrColTimeLine = [
            'employeeId',
            'employeeName',
            'employeeRole',
            'monday',
            'tuesday',
            'wednesday',
            'thursday',
            'friday',
            'saturday',
            'sunday',
            'employeePhone',
            'dateStart',
            'dateEnd',
        ];
        $endSheet = "ca s";
        $time = [];

        for ($row = 1; $row <= $highestRow; $row++) {
            $dataTimeLine = [];
            if ($row <= 5 and $row != 2) {
                continue;
            }

            for ($col = 1; $col <= $highestColumnIndex; ++$col) {
                if ($col == 1) {
                    continue;
                }
                $value = $worksheet->getCellByColumnAndRow($col, $row)->getValue();

                if ($col == 3 and $row == 2) {
                    $explode = explode(" ", $value);
                    $startTime = $explode[38];
                    $endTime = $explode[40];
                    $time['start'] = date('Y-m-d', strtotime($startTime));
                    $time['end'] = date('Y-m-d', strtotime($endTime));

                    continue 2;
                }
                if (empty($value) or in_array(strtolower($value), $ignoreHeader)) {
                    continue;
                }
                if (strtolower($value) == $endSheet) {
                    break 2;
                }

                $dataTimeLine[] = $value;
            }

            foreach ($dataTimeLine as $key => $value) {
                $value = trim($value);
                if (empty($value)) {
                    $dataLog = [
                        'id' => $numberRecord,
                        'file' => $this->entity->get('id'),
                        'description' => 'Dữ liệu cột ' . $arrColTimeLine[$key] . ' bị thiếu!',
                    ];
                    $this->saveLogTimeLine($dataLog);
                    continue 2;
                }

                if (
                    $arrColTimeLine[$key] == 'monday' or
                    $arrColTimeLine[$key] == 'tuesday' or
                    $arrColTimeLine[$key] == 'wednesday' or
                    $arrColTimeLine[$key] == 'thursday' or
                    $arrColTimeLine[$key] == 'friday' or
                    $arrColTimeLine[$key] == 'saturday' or
                    $arrColTimeLine[$key] == 'sunday' and
                    !$timeLineType->constants($arrColTimeLine[$key], $value)
                ) {
                    $dataLog = [
                        'id' => $numberRecord,
                        'file' => $entity->get('id'),
                        'description' => 'Dữ liệu cột ' . $arrColTimeLine[$key] . ' không đúng loại!',
                    ];
                    $this->saveLogTimeLine($dataLog);
                    continue 2;
                }
            }

            $this->createTimeLine($dataTimeLine, $time);
        }
    }

    public function createTimeLine($data, $time)
    {
        $insertQuery = $entityManager
            ->getQueryBuilder()
            ->insert()
            ->into('SomeTable')
            ->columns([
                'employeeId',
                'employeeName',
                'employeeRole',
                'monday',
                'tuesday',
                'wednesday',
                'thursday',
                'friday',
                'saturday',
                'sunday',
                'employeePhone',
                'dateStart',
                'dateEnd'
            ])
            ->values([
                'employeeId' => $data[0],
                'employeeName' => $data[1],
                'employeeRole' => $data[2],
                'monday' => $data[3],
                'tuesday' => $data[4],
                'wednesday' => $data[5],
                'thursday' => $data[6],
                'friday' => $data[7],
                'saturday' => $data[8],
                'sunday' => $data[9],
                'employeePhone' => $data[10],
                'dateStart' => $time['start'],
                'dateEnd' => $time['end'],
            ])
            ->build();

        $entityManager->getQueryExecutor()->execute($insertQuery);
    }

    public function saveLogTimeline($data)
    {
        $insertQuery = $entityManager
            ->getQueryBuilder()
            ->insert()
            ->into('SomeTable')
            ->columns([
                'id',
                'fileId',
                'description',
            ])
            ->values([
                'id' => $data['id'],
                'fileId' => $data['fileId'],
                'description' => $data['description'],
            ])
            ->build();

        $entityManager->getQueryExecutor()->execute($insertQuery);
    }
}

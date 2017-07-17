$config = [
            "title"=>[
                "text"=>$projectName
            ],
            "info"=>["sid"=>$summary->sid],
            "headers"=>[
                [
                    "data"=>"onlinet",
                    "rowspan"=>2,
                    "colspan"=>1,
                    "title"=>"上线时间",
                    "width"=>25
                ],
                [
                    "data"=>"downlinet",
                    "rowspan"=>2,
                    "title"=>"下线时间",
                    "width"=>25
                ],
                [
                    "title"=>"着陆页",
                    "rowspan"=>2,
                    "width"=>30,
                    "render"=>function($model)
                    {
                        $lp = Lp::findOne(["landing"=>$model->landing]);
                        if($lp)
                        {
                            return $lp->name;
                        }else
                        {
                            return $model->landing;
                        }
                    }
                ],
                [
                    "title"=>"成本",
                    "headers"=>[
                        [
                            "title"=>"渠道名",
                            "width"=>30,
                            "render"=>function($model)
                            {
                                $place = Place::findOne(["aid"=>$model->aid]);
                                return $place->name;
                            }
                        ],
                        [
                            "title"=>"点击单价",
                            "width"=>15,
                            "render"=>function($model)
                            {
                                $price = "-";
                                if($model->click_num!=0)
                                {
                                    $price = round($model->cost/$model->click_num,2);
                                }
                                return $price;
                            },
                            "summary"=>["type"=>ExcelHelper::SUMMARY_TEXT,"text"=>"合计:"]
                        ],
                        [
                            "data"=>"click_num",
                            "title"=>"着陆页点击",
                            "width"=>15,
                            "summary"=>["type"=>ExcelHelper::SUMMARY_SUM]
                        ],
                        [
                            "data"=>"cost",
                            "title"=>"着陆页成本",
                            "width"=>15,
                            "summary"=>["type"=>ExcelHelper::SUMMARY_SUM]
                        ]
                    ]
                ]
            ],
            "sheets"=>[
                [
                    "name"=>$projectName,
                    "data"=>$details
                ]
            ]
        ];
        $fileName = $projectName.'('.date("Y年m月j日",strtotime($startDay))."-".date("Y年m月j日",strtotime($endDay)).').xls';
        $excel = new ExcelHelper($config);
        return $excel->renderFile($fileName);

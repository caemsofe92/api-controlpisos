let express = require("express");
let router = express.Router();
const axios = require("axios");
const client = require("../bin/redis-client");
const moment = require("moment");
const { BlobServiceClient } = require("@azure/storage-blob");
require('moment/locale/es');

router.post("/", async (req, res) => {
  try {
    const tenantUrl = req.query.tenantUrl || (req.body && req.body.tenantUrl);
    const clientId = req.query.clientId || (req.body && req.body.clientId);
    const clientSecret =
      req.query.clientSecret || (req.body && req.body.clientSecret);
    const tenant = req.query.tenant || (req.body && req.body.tenant);
    const environment =
      req.query.environment || (req.body && req.body.environment);
    const unsafeCondition =
      req.query.unsafeCondition || (req.body && req.body.unsafeCondition);
    const improvementOpportunity =
      req.query.improvementOpportunity ||
      (req.body && req.body.improvementOpportunity);

    const eventDetails =
      req.query.eventDetails || (req.body && req.body.eventDetails);
    const eventCauses =
      req.query.eventCauses || (req.body && req.body.eventCauses);
    const potentialEventDamage =
      req.query.potentialEventDamage ||
      (req.body && req.body.potentialEventDamage);
    const evidences = req.query.evidences || (req.body && req.body.evidences);
    const email = req.query.email || (req.body && req.body.email);

    if (!tenantUrl || tenantUrl.length === 0)
      throw new Error("tenantUrl is Mandatory");

    if (!clientId || clientId.length === 0)
      throw new Error("clientId is Mandatory");

    if (!clientSecret || clientSecret.length === 0)
      throw new Error("clientSecret is Mandatory");

    if (!tenant || tenant.length === 0) throw new Error("tenant is Mandatory");

    if (!environment || environment.length === 0)
      throw new Error("environment is Mandatory");

    if (!unsafeCondition || unsafeCondition.length === 0)
      throw new Error("unsafeCondition is Mandatory");

    if (!client.isOpen) client.connect();

    let token = await client.get(environment);

    if (!token) {
      const tokenResponse = await axios
        .post(
          `https://login.microsoftonline.com/${tenantUrl}/oauth2/token`,
          `grant_type=client_credentials&client_id=${clientId}&client_secret=${clientSecret}&resource=${tenant}/`,
          { headers: { "Content-Type": "application/x-www-form-urlencoded" } }
        )
        .catch(function (error) {
          if (
            error.response &&
            error.response.data &&
            error.response.data.error &&
            error.response.data.error.innererror &&
            error.response.data.error.innererror.message
          ) {
            throw new Error(error.response.data.error.innererror.message);
          } else if (error.request) {
            throw new Error(error.request);
          } else {
            throw new Error("Error", error.message);
          }
        });
      token = tokenResponse.data.access_token;
      await client.set(environment, tokenResponse.data.access_token, {
        EX: 3599,
      });
    }

    let _unsafeCondition = await axios
      .post(
        `${tenant}/data/UnsafeConditionsReports?$format=application/json;odata.metadata=none`,
        {
          ...unsafeCondition,
          Responsible: unsafeCondition.Responsible.toString(),
          UtcDrawingDate: moment(unsafeCondition.UtcDrawingDate).add(
            5,
            "hours"
          ),
        },
        { headers: { Authorization: "Bearer " + token } }
      )
      .catch(function (error) {
        if (
          error.response &&
          error.response.data &&
          error.response.data.error &&
          error.response.data.error.innererror &&
          error.response.data.error.innererror.message
        ) {
          throw new Error(error.response.data.error.innererror.message);
        } else if (error.request) {
          throw new Error(error.request);
        } else {
          throw new Error("Error", error.message);
        }
      });

    _unsafeCondition = _unsafeCondition.data;

    let _improvementOpportunity;

    if (improvementOpportunity) {
      _improvementOpportunity = await axios
        .post(
          `${tenant}/api/services/SRF_HSEDocuRefServicesGroup/SRF_HSEDocuRefServices/createOpportunities`,
          {
            _description: improvementOpportunity,
            _dataAreaId: _unsafeCondition.dataAreaId,
            _idOrigin: _unsafeCondition.SRF_HSEIdUnsafeCondition,
            _detectionDate: _unsafeCondition.UtcDrawingDate,
            _refRecId: _unsafeCondition.RecId1,
            _state: 0,
            _hcmEmploymentType: 0,
            _origin: 8,
            _tableID: 20371,
          },
          {
            headers: { Authorization: "Bearer " + token },
          }
        )
        .catch(function (error) {
          if (
            error.response &&
            error.response.data &&
            error.response.data.error &&
            error.response.data.error.innererror &&
            error.response.data.error.innererror.message
          ) {
            throw new Error(error.response.data.error.innererror.message);
          } else if (error.request) {
            throw new Error(error.request);
          } else {
            throw new Error("Error", error.message);
          }
        });
      _improvementOpportunity = _improvementOpportunity.data
        ? {
            SRF_HSEIdImprovementOpportunities: _improvementOpportunity.data,
            Description: improvementOpportunity,
            RefRecId: _unsafeCondition.RecId1,
          }
        : {};

      await axios
        .patch(
          `${tenant}/data/UnsafeConditionsReports(dataAreaId='${_unsafeCondition.dataAreaId}',SRF_HSEIdUnsafeCondition='${_unsafeCondition.SRF_HSEIdUnsafeCondition}')?cross-company=true`,
          {
            SRF_HSEIdImprovementOpportunities:
              _improvementOpportunity.SRF_HSEIdImprovementOpportunities,
          },
          {
            headers: { Authorization: "Bearer " + token },
          }
        )
        .catch(function (error) {
          if (
            error.response &&
            error.response.data &&
            error.response.data.error &&
            error.response.data.error.innererror &&
            error.response.data.error.innererror.message
          ) {
            throw new Error(error.response.data.error.innererror.message);
          } else if (error.request) {
            throw new Error(error.request);
          } else {
            throw new Error("Error", error.message);
          }
        });

      _unsafeCondition.SRF_HSEIdImprovementOpportunities =
        _improvementOpportunity.SRF_HSEIdImprovementOpportunities;
    }

    let _eventDetails;

    if (eventDetails) {
      _eventDetails = await axios
        .post(
          `${tenant}/data/SRF_HSEEventDetails?$format=application/json;odata.metadata=none`,
          {
            ...eventDetails,
            dataAreaId: _unsafeCondition.dataAreaId,
            SRF_HSEIdUnsafeCondition: _unsafeCondition.SRF_HSEIdUnsafeCondition,
            EventDate2: moment(eventDetails.EventDate2).add(5, "hours"),
          },
          {
            headers: { Authorization: "Bearer " + token },
          }
        )
        .catch(function (error) {
          if (
            error.response &&
            error.response.data &&
            error.response.data.error &&
            error.response.data.error.innererror &&
            error.response.data.error.innererror.message
          ) {
            throw new Error(error.response.data.error.innererror.message);
          } else if (error.request) {
            throw new Error(error.request);
          } else {
            throw new Error("Error", error.message);
          }
        });
      _eventDetails = _eventDetails.data;
    }

    let _eventCauses = [];

    if (eventCauses) {
      for (let i = 0; i < eventCauses.length; i++) {
        const cause = eventCauses[i];
        const causeResponse = await axios
          .post(
            `${tenant}/data/SRF_HSEEventCauses?$format=application/json;odata.metadata=none`,
            {
              ...cause,
              dataAreaId: _eventDetails.dataAreaId,
              SRF_HSEIdUnsafeCondition: _eventDetails.SRF_HSEIdUnsafeCondition,
              RefRecid: _eventDetails.RecId1,
              Description: undefined,
            },
            {
              headers: { Authorization: "Bearer " + token },
            }
          )
          .catch(function (error) {
            if (
              error.response &&
              error.response.data &&
              error.response.data.error &&
              error.response.data.error.innererror &&
              error.response.data.error.innererror.message
            ) {
              throw new Error(error.response.data.error.innererror.message);
            } else if (error.request) {
              throw new Error(error.request);
            } else {
              throw new Error("Error", error.message);
            }
          });
        _eventCauses.push(causeResponse.data);
      }
    }

    let _potentialEventDamage = [];

    if (potentialEventDamage) {
      for (let i = 0; i < potentialEventDamage.length; i++) {
        const damage = potentialEventDamage[i];
        const damageResponse = await axios
          .post(
            `${tenant}/data/SRF_HSEPotentialEventDamage?$format=application/json;odata.metadata=none`,
            {
              ...damage,
              dataAreaId: _eventDetails.dataAreaId,
              SRF_HSEIdUnsafeCondition: _eventDetails.SRF_HSEIdUnsafeCondition,
              RefRecid: _eventDetails.RecId1,
            },
            {
              headers: { Authorization: "Bearer " + token },
            }
          )
          .catch(function (error) {
            if (
              error.response &&
              error.response.data &&
              error.response.data.error &&
              error.response.data.error.innererror &&
              error.response.data.error.innererror.message
            ) {
              throw new Error(error.response.data.error.innererror.message);
            } else if (error.request) {
              throw new Error(error.request);
            } else {
              throw new Error("Error", error.message);
            }
          });
        _potentialEventDamage.push(damageResponse.data);
      }
    }

    let _evidences = [];

    if (evidences) {
      const blobServiceClient = BlobServiceClient.fromConnectionString(
        process.env.BLOBSTORAGECONNECTIONSTRING
      );

      const containerClient =
        blobServiceClient.getContainerClient(process.env.BLOBSTORAGERAICPATH);

      for (let i = 0; i < evidences.length; i++) {
        const element = evidences[i];

        if (element.imagePath.length > 0) {
          const path = JSON.parse(element.imagePath).toString();

          const matches = path.match(/^data:([A-Za-z-+\/]+);base64,(.+)$/);

          const buffer = new Buffer.from(matches[2], "base64");

          const imageType = matches[1];

          const name =
            _unsafeCondition.RecId1 +
            moment().format().toString() +
            "hseqraicimage." +
            imageType.split("/")[1];

          const blockBlobClient = containerClient.getBlockBlobClient(name);

          const responseImage = await blockBlobClient.upload(
            buffer,
            buffer.byteLength
          );

          const imageRequest = {
            _DataareaId: _unsafeCondition.dataAreaId,
            _AccesInformation: `${process.env.BLOBSTORAGEURL}/${process.env.BLOBSTORAGERAICPATH}/${name}`,
            _name: name,
            _TableId: 20371,
            _RefRecId: _unsafeCondition.RecId1,
            _FileType: imageType.split("/")[1],
          };

          if (responseImage) {
            await axios
              .post(
                `${tenant}/api/services/SRF_HSEDocuRefServicesGroup/SRF_HSEDocuRefServices/FillDocuRef`,
                imageRequest,
                {
                  headers: { Authorization: "Bearer " + token },
                }
              )
              .catch(function (error) {
                if (
                  error.response &&
                  error.response.data &&
                  error.response.data.error &&
                  error.response.data.error.innererror &&
                  error.response.data.error.innererror.message
                ) {
                  throw new Error(error.response.data.error.innererror.message);
                } else if (error.request) {
                  throw new Error(error.request);
                } else {
                  throw new Error("Error", error.message);
                }
              });
            _evidences.push({
              RefRecId: _unsafeCondition.RecId1,
              OriginalFileName: name,
            });
          }
        }
      }
    }

    if (email) {
      await axios
        .post(
          process.env.EMAILNOTIFICATIONURL,
          {
            recipients: !email.recipients || email.recipients === "" ? process.env.DEVELOPEREMAIL : email.recipients,
            message: `<div><p>Señores</p><p>Cordial saludo;</p><p>Nos permitimos notificarles que el ${moment(unsafeCondition.UtcDrawingDate).format('DD/MM/YYYY')} a las ${moment(unsafeCondition.UtcDrawingDate).format('h:mm:ss a')}, ${email.Responsable} ha generado el ${_unsafeCondition.SRF_HSEIdUnsafeCondition} de tipo “${email.TipoReporte}” en ${email.Company} para su respectiva gestión y cierre.</p><p>Gracias</p></div>`,
            subject: `Reporte de actos, incidentes y condiciones inseguras creado - ${_unsafeCondition.SRF_HSEIdUnsafeCondition} ${email.Company}`,
          },
          {
            headers: { "Content-Type": "application/json" },
          }
        )
        .catch(function (error) {
          if (
            error.response &&
            error.response.data &&
            error.response.data.error &&
            error.response.data.error.innererror &&
            error.response.data.error.innererror.message
          ) {
            throw new Error(error.response.data.error.innererror.message);
          } else if (error.request) {
            throw new Error(error.request);
          } else {
            throw new Error("Error", error.message);
          }
        });
    }

    return res.json({
      result: true,
      message: "OK",
      _unsafeCondition,
      _improvementOpportunity,
      _eventDetails,
      _eventCauses,
      _potentialEventDamage,
      _evidences,
    });
  } catch (error) {
    return res.status(500).json({
      result: false,
      message: error.toString(),
    });
  }
});

module.exports = router;
